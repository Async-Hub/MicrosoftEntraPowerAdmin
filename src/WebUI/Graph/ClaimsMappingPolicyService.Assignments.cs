using CSharpFunctionalExtensions;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed partial class ClaimsMappingPolicyService
{
  public Result<ClaimsMappingPolicyAssignmentContext, GraphOperationError> BeginAssignment() => tenants.CurrentTenant.HasValue
    ? new ClaimsMappingPolicyAssignmentContext(tenants.CurrentTenant.Value.TenantId, operation.SelectionId)
    : new GraphOperationError(GraphOperationErrorType.TenantNotSelected, "Select a tenant to manage policy assignments.");

  private bool IsCurrent(ClaimsMappingPolicyAssignmentContext context) => tenants.CurrentTenant.HasValue
    && tenants.CurrentTenant.Value.TenantId == context.TenantId && operation.SelectionId == context.SelectionId;

  private static GraphOperationError AssignmentTenantChanged() => new(GraphOperationErrorType.TenantChanged,
    "The tenant changed. Start a new assignment confirmation. If already submitted, check assignments in the original tenant.");

  public Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> GetAssignedPoliciesAsync(
    ClaimsMappingPolicyAssignmentContext context, Guid servicePrincipalObjectId, CancellationToken cancellationToken = default) =>
    RunAssignmentAsync(context, client => ReadAssignedPoliciesAsync(client, context, servicePrincipalObjectId, cancellationToken), cancellationToken);

  private async Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> ReadAssignedPoliciesAsync(
    GraphServiceClient client, ClaimsMappingPolicyAssignmentContext context, Guid servicePrincipalObjectId, CancellationToken cancellationToken)
  {
    if (servicePrincipalObjectId == Guid.Empty)
      return DirectoryReadOperation.InvalidId();
    var builder = client.ServicePrincipals[servicePrincipalObjectId.ToString("D")].ClaimsMappingPolicies;
    var path = $"servicePrincipals/{servicePrincipalObjectId:D}/claimsMappingPolicies";
    var items = new List<ClaimsMappingPolicyListItem>();
    var seenPages = new HashSet<string>(StringComparer.Ordinal);
    string? next = null;
    do
    {
      cancellationToken.ThrowIfCancellationRequested();
      if (!IsCurrent(context))
        return AssignmentTenantChanged();
      var response = next is null
        ? await builder.GetAsync(request =>
        {
          request.QueryParameters.Select = PolicyFields;
          request.QueryParameters.Top = 100;
        }, cancellationToken)
        : await builder.WithUrl(next).GetAsync(cancellationToken: cancellationToken);
      if (!IsCurrent(context))
        return AssignmentTenantChanged();
      if (response?.Value is not { } page || page.Any(item => !HasValidId(item)))
        return InvalidResponse();
      items.AddRange(page.Select(MapPolicy));
      var continuation = operation.Continuation(response.OdataNextLink, path);
      if (continuation.IsFailure)
        return continuation.Error;
      next = continuation.Value?.Url;
      if (next is not null && !seenPages.Add(next))
        return InvalidResponse();
    } while (next is not null);
    return items.AsReadOnly();
  }

  public Task<UnitResult<GraphOperationError>> AssignAsync(ClaimsMappingPolicyAssignmentContext context,
    Guid servicePrincipalObjectId, Guid claimsMappingPolicyObjectId, CancellationToken cancellationToken = default) =>
    ChangeAssignmentAsync(context, servicePrincipalObjectId, claimsMappingPolicyObjectId, unassign: false, cancellationToken);

  public Task<UnitResult<GraphOperationError>> UnassignAsync(ClaimsMappingPolicyAssignmentContext context,
    Guid servicePrincipalObjectId, Guid claimsMappingPolicyObjectId, CancellationToken cancellationToken = default) =>
    ChangeAssignmentAsync(context, servicePrincipalObjectId, claimsMappingPolicyObjectId, unassign: true, cancellationToken);

  private async Task<UnitResult<GraphOperationError>> ChangeAssignmentAsync(ClaimsMappingPolicyAssignmentContext context,
    Guid servicePrincipalObjectId, Guid claimsMappingPolicyObjectId, bool unassign, CancellationToken cancellationToken)
  {
    if (servicePrincipalObjectId == Guid.Empty || claimsMappingPolicyObjectId == Guid.Empty)
      return DirectoryReadOperation.InvalidId();
    var result = await RunAssignmentAsync<bool>(context, async client =>
    {
      // Resolve both tenant-specific objects on the same pinned client before changing their relationship.
      var principal = await client.ServicePrincipals[servicePrincipalObjectId.ToString("D")].GetAsync(
        request => request.QueryParameters.Select = ["id"], cancellationToken);
      if (!IsCurrent(context))
        return AssignmentTenantChanged();
      if (principal is null)
        return AssignmentNotFound();
      if (!HasValidId(principal) || Guid.Parse(principal.Id!) != servicePrincipalObjectId)
        return InvalidResponse();
      var policy = await client.Policies.ClaimsMappingPolicies[claimsMappingPolicyObjectId.ToString("D")].GetAsync(
        request => request.QueryParameters.Select = ["id"], cancellationToken);
      if (!IsCurrent(context))
        return AssignmentTenantChanged();
      if (policy is null)
        return AssignmentNotFound();
      if (!HasValidId(policy) || Guid.Parse(policy.Id!) != claimsMappingPolicyObjectId)
        return InvalidResponse();
      var assigned = await ReadAssignedPoliciesAsync(client, context, servicePrincipalObjectId, cancellationToken);
      if (assigned.IsFailure)
        return assigned.Error;
      var exists = assigned.Value.Any(item => item.ObjectId == claimsMappingPolicyObjectId);
      if (exists != unassign)
        return new GraphOperationError(GraphOperationErrorType.InvalidInput, exists
          ? "This policy is already assigned. Assignment state has been refreshed."
          : "This relationship is already unassigned. Assignment state has been refreshed.");
      if (!IsCurrent(context))
        return AssignmentTenantChanged();
      cancellationToken.ThrowIfCancellationRequested();
      var relationship = client.ServicePrincipals[servicePrincipalObjectId.ToString("D")].ClaimsMappingPolicies;
      if (unassign)
        await relationship[claimsMappingPolicyObjectId.ToString("D")].Ref.DeleteAsync(cancellationToken: cancellationToken);
      else
        await relationship.Ref.PostAsync(new ReferenceCreate
        {
          OdataId = $"https://graph.microsoft.com/v1.0/policies/claimsMappingPolicies/{claimsMappingPolicyObjectId:D}"
        }, cancellationToken: cancellationToken);
      return true;
    }, cancellationToken);
    return result.IsSuccess ? UnitResult.Success<GraphOperationError>() : UnitResult.Failure(result.Error);
  }

  private async Task<Result<T, GraphOperationError>> RunAssignmentAsync<T>(ClaimsMappingPolicyAssignmentContext context,
    Func<GraphServiceClient, Task<Result<T, GraphOperationError>>> execute, CancellationToken cancellationToken)
  {
    cancellationToken.ThrowIfCancellationRequested();
    if (!IsCurrent(context))
      return AssignmentTenantChanged();
    try
    {
      var created = await clients.CreateAsync(cancellationToken);
      if (created.IsFailure)
        return IsCurrent(context) ? AssignmentFailure(created.Error) : AssignmentTenantChanged();
      using var client = created.Value;
      if (!IsCurrent(context))
        return AssignmentTenantChanged();
      var result = await execute(client);
      cancellationToken.ThrowIfCancellationRequested();
      return IsCurrent(context) ? result : AssignmentTenantChanged();
    }
    catch (MicrosoftIdentityWebChallengeUserException exception)
    {
      return AssignmentFailure(GraphErrorMapping.From(exception.MsalUiRequiredException));
    }
    catch (MsalUiRequiredException exception)
    {
      return AssignmentFailure(GraphErrorMapping.From(exception));
    }
    catch (MsalException exception)
    {
      return AssignmentFailure(GraphErrorMapping.From(exception));
    }
    catch (ApiException exception)
    {
      var error = GraphErrorMapping.From(exception);
      return AssignmentFailure(error with
      {
        Type = exception.ResponseStatusCode switch
        {
          404 => GraphOperationErrorType.NotFound,
          400 or 409 or 422 => GraphOperationErrorType.InvalidInput,
          _ => error.Type
        },
        Message = exception.ResponseStatusCode switch
        {
          404 => AssignmentNotFound().Message,
          400 or 409 or 422 => "Microsoft Graph rejected this relationship change. It may already exist or have changed. Refresh assignments before trying again.",
          _ when error.Type == GraphOperationErrorType.GraphFailure => "Microsoft Graph could not confirm assignment state. Refresh assignments before trying again.",
          _ => error.Message
        }
      });
    }
    catch (HttpRequestException)
    {
      return AssignmentFailure(new(GraphOperationErrorType.GraphFailure,
        "Microsoft Graph could not be reached. A submitted change may have completed; refresh assignments before trying again."));
    }
    catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
    {
      return AssignmentFailure(new(GraphOperationErrorType.GraphFailure,
        "Microsoft Graph timed out. A submitted change may have completed; refresh assignments before trying again."));
    }
  }

  private static GraphOperationError AssignmentNotFound() => new(GraphOperationErrorType.NotFound,
    "The Service Principal, policy, or relationship was not found in the current tenant. Refresh assignments.");

  private GraphOperationError AssignmentFailure(GraphOperationError error)
  {
    if (error.Type is GraphOperationErrorType.Forbidden or GraphOperationErrorType.ConsentRequired)
      error = error with { Message = "MEPA does not currently have permission to manage Enterprise Application policy assignments in this tenant. Check delegated Application.ReadWrite.All and Policy.ReadWrite.ApplicationConfiguration admin consent and your administrator privileges. Authenticate the tenant after consent is granted." };
    logger.LogWarning("Claims mapping policy assignment operation failed: {ErrorType}, HTTP {Status}, code {Code}, request {RequestId}",
      error.Type, error.HttpStatus, error.Code, error.RequestId);
    return error;
  }
}
