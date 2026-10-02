using CSharpFunctionalExtensions;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed partial class ClaimsMappingPolicyService
{
  public async Task<Result<ClaimsMappingPolicyDeletionContext, GraphOperationError>> BeginDeleteAsync(
    Guid objectId, CancellationToken cancellationToken = default)
  {
    if (tenants.CurrentTenant.HasNoValue)
      return new GraphOperationError(GraphOperationErrorType.TenantNotSelected, "Select a tenant to delete a policy.");
    if (objectId == Guid.Empty)
      return DirectoryReadOperation.InvalidId();
    var tenantId = tenants.CurrentTenant.Value.TenantId;
    var selection = operation.SelectionId;
    return await RunDeletionAsync<ClaimsMappingPolicyDeletionContext>(tenantId, selection, objectId, async client =>
    {
      var policy = await client.Policies.ClaimsMappingPolicies[objectId.ToString("D")].GetAsync(
        request => request.QueryParameters.Select = PolicyFields, cancellationToken);
      if (!IsCurrentDeletion(tenantId, selection))
        return DeleteTenantChanged();
      if (policy is null)
        return DeleteNotFound();
      if (!HasValidId(policy) || Guid.Parse(policy.Id!) != objectId)
        return InvalidResponse();
      var assignments = await ReadAssignmentsAsync(client, objectId, selection, cancellationToken);
      if (assignments.IsFailure)
        return assignments.Error;
      return new ClaimsMappingPolicyDeletionContext(tenantId, selection,
        new(MapPolicy(policy), assignments.Map(MapAssignments)), policy.DisplayName);
    }, cancellationToken);
  }

  public async Task<UnitResult<GraphOperationError>> DeleteAsync(ClaimsMappingPolicyDeletionContext context,
    string confirmation, CancellationToken cancellationToken = default)
  {
    cancellationToken.ThrowIfCancellationRequested();
    if (!IsCurrentDeletion(context.TenantId, context.SelectionId))
      return UnitResult.Failure(DeleteTenantChanged());
    if (context.Original.Policy.ObjectId == Guid.Empty)
      return UnitResult.Failure(DirectoryReadOperation.InvalidId());
    if (!context.CanConfirm(confirmation))
      return UnitResult.Failure(context.HasAssignments ? DeleteAssigned() : new GraphOperationError(
        GraphOperationErrorType.ConfirmationRequired, "Review the current assignments and type the policy name to confirm deletion."));

    var objectId = context.Original.Policy.ObjectId;
    logger.LogInformation("Claims mapping policy deletion requested in tenant {TenantId} for policy {PolicyId}", context.TenantId, objectId);
    var result = await RunDeletionAsync<bool>(context.TenantId, context.SelectionId, objectId, async client =>
    {
      var builder = client.Policies.ClaimsMappingPolicies[objectId.ToString("D")];
      var policy = await builder.GetAsync(request => request.QueryParameters.Select = ["id", "displayName"], cancellationToken);
      if (!IsCurrentDeletion(context.TenantId, context.SelectionId))
        return DeleteTenantChanged();
      if (policy is null)
        return DeleteNotFound();
      if (!HasValidId(policy) || Guid.Parse(policy.Id!) != objectId)
        return InvalidResponse();
      if (!string.Equals(policy.DisplayName, context.CurrentDisplayName, StringComparison.Ordinal))
        return new GraphOperationError(GraphOperationErrorType.StaleState,
          "The policy name changed. Reopen deletion and confirm the current name before trying again.");
      if (!ClaimsMappingPolicyDeletionContext.Matches(policy.DisplayName, confirmation))
        return new GraphOperationError(GraphOperationErrorType.ConfirmationRequired, "Type the current policy name to confirm deletion.");

      // Reuse the fully paged appliesTo read on this tenant-pinned client immediately before DELETE.
      var assignments = await ReadAssignmentsAsync(client, objectId, context.SelectionId, cancellationToken);
      if (!IsCurrentDeletion(context.TenantId, context.SelectionId))
        return DeleteTenantChanged();
      if (assignments.IsFailure)
        return assignments.Error;
      if (assignments.Value.Count > 0)
        return DeleteAssigned();
      cancellationToken.ThrowIfCancellationRequested();
      await builder.DeleteAsync(cancellationToken: cancellationToken);
      logger.LogInformation("Claims mapping policy deleted in tenant {TenantId}: {PolicyId}", context.TenantId, objectId);
      return true;
    }, cancellationToken);
    return result.IsSuccess ? UnitResult.Success<GraphOperationError>() : UnitResult.Failure(result.Error);
  }

  private bool IsCurrentDeletion(Guid tenantId, Guid selection) => tenants.CurrentTenant.HasValue
    && tenants.CurrentTenant.Value.TenantId == tenantId && operation.SelectionId == selection;

  private static GraphOperationError DeleteTenantChanged() => new(GraphOperationErrorType.TenantChanged,
    "The tenant changed. Reopen the policy in the selected tenant. If deletion was already submitted, check the original tenant's policy list.");
  private static GraphOperationError DeleteNotFound() => new(GraphOperationErrorType.NotFound,
    "This policy no longer exists in the selected tenant. Refresh the policy list.");
  private static GraphOperationError DeleteAssigned() => new(GraphOperationErrorType.PolicyIsAssigned,
    "This policy is now assigned to one or more directory objects and cannot be deleted. Unassign it from all Enterprise Applications before trying again.");

  private async Task<Result<T, GraphOperationError>> RunDeletionAsync<T>(Guid tenantId, Guid selection, Guid objectId,
    Func<GraphServiceClient, Task<Result<T, GraphOperationError>>> execute, CancellationToken cancellationToken)
  {
    cancellationToken.ThrowIfCancellationRequested();
    if (!IsCurrentDeletion(tenantId, selection))
      return DeleteTenantChanged();
    try
    {
      var created = await clients.CreateAsync(cancellationToken);
      if (created.IsFailure)
        return Failure(created.Error);
      using var client = created.Value;
      if (!IsCurrentDeletion(tenantId, selection))
        return DeleteTenantChanged();
      var result = await execute(client);
      cancellationToken.ThrowIfCancellationRequested();
      if (!IsCurrentDeletion(tenantId, selection))
        return DeleteTenantChanged();
      return result.IsFailure ? Failure(result.Error) : result;
    }
    catch (MicrosoftIdentityWebChallengeUserException exception)
    {
      return Failure(GraphErrorMapping.From(exception.MsalUiRequiredException));
    }
    catch (MsalUiRequiredException exception)
    {
      return Failure(GraphErrorMapping.From(exception));
    }
    catch (MsalException exception)
    {
      return Failure(GraphErrorMapping.From(exception));
    }
    catch (ApiException exception)
    {
      var error = GraphErrorMapping.From(exception);
      return Failure(error with
      {
        Type = exception.ResponseStatusCode == 404 ? GraphOperationErrorType.NotFound : error.Type,
        Message = exception.ResponseStatusCode switch
        {
          404 => DeleteNotFound().Message,
          _ when error.Type == GraphOperationErrorType.GraphFailure => "Microsoft Graph could not confirm deletion. Refresh the policy list before trying again.",
          _ => error.Message
        }
      });
    }
    catch (HttpRequestException)
    {
      return Failure(new(GraphOperationErrorType.GraphFailure,
        "Microsoft Graph could not be reached. A submitted deletion may have completed; refresh the policy list before trying again."));
    }
    catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
    {
      return Failure(new(GraphOperationErrorType.GraphFailure,
        "Microsoft Graph timed out. A submitted deletion may have completed; refresh the policy list before trying again."));
    }

    GraphOperationError Failure(GraphOperationError error)
    {
      if (!IsCurrentDeletion(tenantId, selection))
        error = DeleteTenantChanged();
      if (error.Type is GraphOperationErrorType.Forbidden or GraphOperationErrorType.ConsentRequired)
        error = error with { Message = "Policy deletion requires delegated Policy.ReadWrite.ApplicationConfiguration admin consent and appropriate administrator privileges in this tenant. Assignment discovery also requires the existing Policy.Read.All and Application.Read.All permissions. Authenticate the tenant after consent is granted." };
      logger.LogWarning("Claims mapping policy deletion workflow failed in tenant {TenantId} for policy {PolicyId}: {ErrorType}, HTTP {Status}, code {Code}, request {RequestId}",
        tenantId, objectId, error.Type, error.HttpStatus, error.Code, error.RequestId);
      return error;
    }
  }
}
