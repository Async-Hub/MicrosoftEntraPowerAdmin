using CSharpFunctionalExtensions;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using GraphPolicy = Microsoft.Graph.Models.ClaimsMappingPolicy;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed class ClaimsMappingPolicyService(DirectoryReadOperation operation) : IClaimsMappingPolicyService
{
  private const string Collection = "policies/claimsMappingPolicies";
  private static readonly string[] PolicyFields = ["id", "displayName", "definition", "isOrganizationDefault"];

  public Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> ListAsync(
    CancellationToken cancellationToken = default) =>
    operation.RunAsync<IReadOnlyList<ClaimsMappingPolicyListItem>>(async client =>
    {
      var policies = await ReadPoliciesAsync(client, operation.SelectionId, cancellationToken);
      return policies.Map(items => (IReadOnlyList<ClaimsMappingPolicyListItem>)Array.AsReadOnly(items.Select(MapPolicy).ToArray()));
    }, cancellationToken, policyRead: true);

  public Task<Result<ClaimsMappingPolicyDetails, GraphOperationError>> GetByObjectIdAsync(
    Guid objectId, CancellationToken cancellationToken = default) =>
    operation.RunAsync<ClaimsMappingPolicyDetails>(async client =>
    {
      if (objectId == Guid.Empty)
        return DirectoryReadOperation.InvalidId();
      var selection = operation.SelectionId;
      var policy = await client.Policies.ClaimsMappingPolicies[objectId.ToString("D")].GetAsync(request =>
        request.QueryParameters.Select = PolicyFields, cancellationToken);
      if (policy is null)
        return new GraphOperationError(GraphOperationErrorType.NotFound, "The policy was not found in the current tenant. Refresh the policy list.");
      if (!HasValidId(policy) || Guid.Parse(policy.Id!) != objectId)
        return InvalidResponse();
      if (selection != operation.SelectionId)
        return DirectoryReadOperation.TenantChanged();

      // A separate boundary lets details remain inspectable when assignment discovery fails.
      // It still rejects a tenant switch before creating a client or using this policy ID.
      var assignments = await operation.RunAsync<ClaimsMappingAssignments>(async assignmentClient =>
      {
        if (selection != operation.SelectionId)
          return DirectoryReadOperation.TenantChanged();
        var objects = await ReadAssignmentsAsync(assignmentClient, objectId, selection, cancellationToken);
        return objects.Map(MapAssignments);
      }, cancellationToken, policyRead: true);
      return new ClaimsMappingPolicyDetails(MapPolicy(policy), assignments);
    }, cancellationToken, policyRead: true);

  public Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> FindForServicePrincipalAsync(
    Guid servicePrincipalObjectId, CancellationToken cancellationToken = default) =>
    operation.RunAsync<IReadOnlyList<ClaimsMappingPolicyListItem>>(async client =>
    {
      if (servicePrincipalObjectId == Guid.Empty)
        return DirectoryReadOperation.InvalidId();
      var selection = operation.SelectionId;
      // The SP navigation endpoint requires a write permission. Policy-side reads
      // use only Policy.Read.All + Application.Read.All. Directory-object expansion
      // can truncate relationships without next links, so page appliesTo separately.
      // https://learn.microsoft.com/en-us/graph/known-issues#some-limitations-apply-to-query-parameters
      var policies = await ReadPoliciesAsync(client, selection, cancellationToken);
      if (policies.IsFailure)
        return policies.Error;
      var matches = new List<ClaimsMappingPolicyListItem>();
      foreach (var policy in policies.Value)
      {
        cancellationToken.ThrowIfCancellationRequested();
        if (selection != operation.SelectionId)
          return DirectoryReadOperation.TenantChanged();
        var assignments = await ReadAssignmentsAsync(client, Guid.Parse(policy.Id!), selection, cancellationToken);
        if (assignments.IsFailure)
          return assignments.Error;
        if (assignments.Value.Any(item => item is ServicePrincipal && Guid.TryParse(item.Id, out var id) && id == servicePrincipalObjectId))
          matches.Add(MapPolicy(policy));
      }
      return matches.AsReadOnly();
    }, cancellationToken, policyRead: true);

  private async Task<Result<IReadOnlyList<GraphPolicy>, GraphOperationError>> ReadPoliciesAsync(
    GraphServiceClient client, Guid selection, CancellationToken cancellationToken)
  {
    var items = new List<GraphPolicy>();
    var seenPages = new HashSet<string>(StringComparer.Ordinal);
    string? next = null;
    do
    {
      cancellationToken.ThrowIfCancellationRequested();
      if (selection != operation.SelectionId)
        return DirectoryReadOperation.TenantChanged();
      var response = next is null
        ? await client.Policies.ClaimsMappingPolicies.GetAsync(request =>
        {
          request.QueryParameters.Select = PolicyFields;
          request.QueryParameters.Top = 100;
        }, cancellationToken)
        : await client.Policies.ClaimsMappingPolicies.WithUrl(next).GetAsync(cancellationToken: cancellationToken);
      if (response?.Value is not { } page || page.Any(item => !HasValidId(item)))
        return InvalidResponse();
      items.AddRange(page);
      var continuation = operation.Continuation(response.OdataNextLink, Collection);
      if (continuation.IsFailure)
        return continuation.Error;
      next = continuation.Value?.Url;
      if (next is not null && !seenPages.Add(next))
        return InvalidResponse();
    } while (next is not null);
    return items.AsReadOnly();
  }

  private async Task<Result<IReadOnlyList<DirectoryObject>, GraphOperationError>> ReadAssignmentsAsync(
    GraphServiceClient client, Guid policyId, Guid selection, CancellationToken cancellationToken)
  {
    var items = new List<DirectoryObject>();
    var seenPages = new HashSet<string>(StringComparer.Ordinal);
    var path = $"{Collection}/{policyId:D}/appliesTo";
    var builder = client.Policies.ClaimsMappingPolicies[policyId.ToString("D")].AppliesTo;
    string? next = null;
    do
    {
      cancellationToken.ThrowIfCancellationRequested();
      if (selection != operation.SelectionId)
        return DirectoryReadOperation.TenantChanged();
      var response = next is null
        ? await builder.GetAsync(request =>
        {
          request.QueryParameters.Select = ["id", "appId", "displayName"];
          request.QueryParameters.Top = 100;
        }, cancellationToken)
        : await builder.WithUrl(next).GetAsync(cancellationToken: cancellationToken);
      if (response?.Value is not { } page || page.Any(item => !HasValidId(item)))
        return InvalidResponse();
      items.AddRange(page);
      var continuation = operation.Continuation(response.OdataNextLink, path);
      if (continuation.IsFailure)
        return continuation.Error;
      next = continuation.Value?.Url;
      if (next is not null && !seenPages.Add(next))
        return InvalidResponse();
    } while (next is not null);
    return items.AsReadOnly();
  }

  private static bool HasValidId(DirectoryObject model) => Guid.TryParse(model.Id, out var id) && id != Guid.Empty;
  private static GraphOperationError InvalidResponse() => new(GraphOperationErrorType.GraphFailure,
    "Microsoft Graph returned invalid claims mapping policy or assignment data. Refresh and try again.");

  private static ClaimsMappingPolicyListItem MapPolicy(GraphPolicy model) => new(
    Guid.Parse(model.Id!), string.IsNullOrWhiteSpace(model.DisplayName) ? "Name unavailable" : model.DisplayName,
    model.IsOrganizationDefault,
    Array.AsReadOnly(model.Definition?.Select(ClaimsMappingDefinitionParser.Parse).ToArray() ?? []));

  private static ClaimsMappingAssignments MapAssignments(IReadOnlyList<DirectoryObject> objects) => new(
    Array.AsReadOnly(objects.OfType<ServicePrincipal>().Select(ServicePrincipalService.MapListItem).ToArray()),
    Array.AsReadOnly(objects.OfType<Application>().Select(ApplicationService.MapListItem).ToArray()),
    Array.AsReadOnly(objects.Where(item => item is not ServicePrincipal and not Application)
      .Select(item => new PolicyDirectoryObject(Guid.Parse(item.Id!), item.OdataType)).ToArray()));
}
