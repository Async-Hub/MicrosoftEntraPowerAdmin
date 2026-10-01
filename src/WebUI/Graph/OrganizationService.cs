using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed class OrganizationService(
  ICurrentTenantContext tenantContext,
  IEntraGraphClientFactory clientFactory,
  ILogger<OrganizationService> logger) : IOrganizationService
{
  public async Task<Result<MicrosoftEntraOrganization, GraphOperationError>> GetCurrentAsync(
    CancellationToken cancellationToken = default)
  {
    var tenant = tenantContext.CurrentTenant;
    if (tenant.HasNoValue)
      return new GraphOperationError(GraphOperationErrorType.TenantNotSelected, "Select a tenant to connect to Microsoft Graph.");

    var tenantId = tenant.Value.TenantId;
    try
    {
      var created = await clientFactory.CreateAsync(cancellationToken);
      if (created.IsFailure)
        return created.Error;

      using var client = created.Value;
      if (tenantContext.CurrentTenant != tenant)
        return TenantChanged();

      logger.LogDebug("Microsoft Graph organization query started for {TenantId}", tenantId);
      var response = await client.Organization.GetAsync(request =>
        request.QueryParameters.Select = ["id", "displayName", "verifiedDomains"], cancellationToken);

      cancellationToken.ThrowIfCancellationRequested();
      if (tenantContext.CurrentTenant != tenant)
        return TenantChanged();

      if (response?.Value is not { Count: 1 } organizations || !string.IsNullOrEmpty(response.OdataNextLink))
        return Failure(new(GraphOperationErrorType.GraphFailure, "Microsoft Graph must return exactly one organization."), tenantId);

      var organization = organizations[0];
      if (!Guid.TryParse(organization.Id, out var actualTenantId) || actualTenantId == Guid.Empty)
        return Failure(new(GraphOperationErrorType.GraphFailure, "Microsoft Graph returned an organization without a valid tenant ID."), tenantId);

      if (actualTenantId != tenantId)
      {
        logger.LogError("Microsoft Graph tenant mismatch: expected {ExpectedTenantId}, received {ActualTenantId}", tenantId, actualTenantId);
        return new GraphOperationError(GraphOperationErrorType.TenantMismatch,
          $"Microsoft Graph returned organization '{actualTenantId}' while tenant '{tenantId}' is selected.");
      }

      var domains = organization.VerifiedDomains?
        .Select(domain => domain.Name).OfType<string>().Where(name => !string.IsNullOrWhiteSpace(name))
        .Distinct(StringComparer.OrdinalIgnoreCase).ToArray() ?? [];
      logger.LogInformation("Microsoft Graph tenant verification succeeded for {TenantId}", tenantId);
      return new MicrosoftEntraOrganization(tenantId,
        string.IsNullOrWhiteSpace(organization.DisplayName) ? "Name unavailable" : organization.DisplayName,
        Array.AsReadOnly(domains));
    }
    catch (MicrosoftIdentityWebChallengeUserException exception)
    {
      return Failure(GraphErrorMapping.From(exception.MsalUiRequiredException), tenantId);
    }
    catch (MsalUiRequiredException exception)
    {
      return Failure(GraphErrorMapping.From(exception), tenantId);
    }
    catch (MsalException exception)
    {
      return Failure(GraphErrorMapping.From(exception), tenantId);
    }
    catch (ApiException exception)
    {
      return Failure(GraphErrorMapping.From(exception), tenantId);
    }
    catch (HttpRequestException)
    {
      return Failure(new(GraphOperationErrorType.GraphFailure, "Microsoft Graph could not be reached. Try again later."), tenantId);
    }
    catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
    {
      return Failure(new(GraphOperationErrorType.GraphFailure, "Microsoft Graph timed out. Try again later."), tenantId);
    }
  }

  private GraphOperationError Failure(GraphOperationError error, Guid tenantId)
  {
    // Exception messages/objects may contain identity payloads; log only structured metadata.
    logger.LogWarning("Microsoft Graph failed for {TenantId}: {ErrorType}, HTTP {Status}, code {Code}, request {RequestId}, client request {ClientRequestId}",
      tenantId, error.Type, error.HttpStatus, error.Code, error.RequestId, error.ClientRequestId);
    return error;
  }

  private static GraphOperationError TenantChanged() => new(
    GraphOperationErrorType.TenantChanged, "The selected tenant changed. Connect again.");
}
