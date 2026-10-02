using CSharpFunctionalExtensions;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed partial class ClaimsMappingPolicyService
{
  public async Task<Result<ClaimsMappingPolicyEditContext, GraphOperationError>> BeginEditAsync(
    Guid objectId, CancellationToken cancellationToken = default)
  {
    if (tenants.CurrentTenant.HasNoValue)
      return new GraphOperationError(GraphOperationErrorType.TenantNotSelected, "Select a tenant to edit a policy.");
    var tenantId = tenants.CurrentTenant.Value.TenantId;
    var selection = operation.SelectionId;
    var result = await GetByObjectIdAsync(objectId, cancellationToken);
    if (selection != operation.SelectionId)
      return EditTenantChanged();
    return result.Map(details => new ClaimsMappingPolicyEditContext(tenantId, selection, details));
  }

  private bool IsCurrent(ClaimsMappingPolicyEditContext context) => tenants.CurrentTenant.HasValue
    && tenants.CurrentTenant.Value.TenantId == context.TenantId && operation.SelectionId == context.SelectionId;

  private static GraphOperationError EditTenantChanged() => new(GraphOperationErrorType.TenantChanged,
    "The tenant changed. Reopen the policy in the selected tenant. If a save was already submitted, check the original tenant.");

  public async Task<Result<ClaimsMappingPolicyDetails, GraphOperationError>> UpdateAsync(
    ClaimsMappingPolicyEditContext context, UpdateClaimsMappingPolicyRequest request, CancellationToken cancellationToken = default)
  {
    cancellationToken.ThrowIfCancellationRequested();
    if (!IsCurrent(context))
      return EditTenantChanged();
    var original = context.Original.Policy;
    if (original.ObjectId == Guid.Empty)
      return DirectoryReadOperation.InvalidId();
    var patch = request.CreatePatch(original);
    if (patch.IsFailure)
      return new GraphOperationError(GraphOperationErrorType.InvalidInput, patch.Error);
    if (patch.Value.HasNoValue)
      return context.Original;

    var submitted = false;
    try
    {
      var created = await clients.CreateAsync(cancellationToken);
      if (created.IsFailure)
        return !IsCurrent(context) ? EditTenantChanged() : EditingFailure(created.Error);
      using var client = created.Value;
      if (!IsCurrent(context))
        return EditTenantChanged();
      var builder = client.Policies.ClaimsMappingPolicies[original.ObjectId.ToString("D")];
      if (patch.Value.Value.Definition is not null)
      {
        // Exact string comparison is deliberately conservative; even formatting changes require a reload.
        var current = await builder.GetAsync(configuration => configuration.QueryParameters.Select = PolicyFields, cancellationToken);
        if (!IsCurrent(context))
          return EditTenantChanged();
        if (current is null)
          return EditNotFound();
        if (!HasValidId(current) || Guid.Parse(current.Id!) != original.ObjectId)
          return InvalidResponse();
        if (!(current.Definition ?? []).SequenceEqual(original.Definitions.Select(definition => definition.Raw), StringComparer.Ordinal))
          return new GraphOperationError(GraphOperationErrorType.StaleState,
            "This policy changed after you opened the editor. Reload the latest version before applying your changes.");
        var assignments = await ReadAssignmentsAsync(client, original.ObjectId, context.SelectionId, cancellationToken);
        if (!IsCurrent(context))
          return EditTenantChanged();
        if (assignments.IsFailure)
          return assignments.Error;
        var ids = assignments.Value.Select(item => Guid.Parse(item.Id!)).Order().ToArray();
        if (ids.Length > 0 && !(request.ConfirmedAssignmentIds ?? []).Order().SequenceEqual(ids))
          return new GraphOperationError(GraphOperationErrorType.ConfirmationRequired,
            "Assignments require confirmation or have changed. Reload the editor and confirm the current assignment impact before saving.");
      }
      if (!IsCurrent(context))
        return EditTenantChanged();
      cancellationToken.ThrowIfCancellationRequested();
      // This is the client pinned to the draft tenant. PATCH preserves the policy ID and relationships.
      await builder.PatchAsync(patch.Value.Value, cancellationToken: cancellationToken);
      submitted = true;
      if (!IsCurrent(context))
        return EditTenantChanged();
      // Graph returns 204: the PATCH response is not an updated policy. Always read authoritative state.
      var reloaded = await GetByObjectIdAsync(original.ObjectId, cancellationToken);
      if (!IsCurrent(context))
        return EditTenantChanged();
      return reloaded.IsSuccess ? reloaded : reloaded.Error with
      {
        Message = "The policy was updated, but its latest state could not be loaded. Reload before editing again. " + reloaded.Error.Message
      };
    }
    catch (MicrosoftIdentityWebChallengeUserException exception)
    {
      return EditingFailure(GraphErrorMapping.From(exception.MsalUiRequiredException));
    }
    catch (MsalUiRequiredException exception)
    {
      return EditingFailure(GraphErrorMapping.From(exception));
    }
    catch (MsalException exception)
    {
      return EditingFailure(GraphErrorMapping.From(exception));
    }
    catch (ApiException exception)
    {
      var error = GraphErrorMapping.From(exception);
      return EditingFailure(error with
      {
        Type = exception.ResponseStatusCode switch
        {
          404 => GraphOperationErrorType.NotFound,
          400 or 409 or 422 => GraphOperationErrorType.InvalidInput,
          _ => error.Type
        },
        Message = exception.ResponseStatusCode switch
        {
          404 => EditNotFound().Message,
          400 or 409 or 422 => "Microsoft Graph rejected the update. Check the name, attributes, and output claim types.",
          _ when error.Type == GraphOperationErrorType.GraphFailure => "Microsoft Graph could not confirm the update. Reload before trying again.",
          _ => error.Message
        }
      });
    }
    catch (HttpRequestException)
    {
      return EditingFailure(new(GraphOperationErrorType.GraphFailure,
        "Microsoft Graph could not be reached. A submitted update may have completed; reload before trying again."));
    }
    catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
    {
      return EditingFailure(new(GraphOperationErrorType.GraphFailure,
        submitted ? "The policy was updated, but reloading timed out. Reload before editing again."
          : "Microsoft Graph timed out. A submitted update may have completed; reload before trying again."));
    }
  }

  private static GraphOperationError EditNotFound() => new(GraphOperationErrorType.NotFound,
    "This policy no longer exists in the selected tenant. Refresh the policy list.");

  private GraphOperationError EditingFailure(GraphOperationError error)
  {
    if (error.Type is GraphOperationErrorType.Forbidden or GraphOperationErrorType.ConsentRequired)
      error = error with { Message = "Policy editing requires delegated Policy.ReadWrite.ApplicationConfiguration admin consent and appropriate directory privileges in this tenant. Authenticate the tenant after consent is granted." };
    logger.LogWarning("Claims mapping policy update failed: {ErrorType}, HTTP {Status}, code {Code}, request {RequestId}",
      error.Type, error.HttpStatus, error.Code, error.RequestId);
    return error;
  }
}
