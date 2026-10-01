using CSharpFunctionalExtensions;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions;
using GraphPolicy = Microsoft.Graph.Models.ClaimsMappingPolicy;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed class ClaimsMappingPolicyService(DirectoryReadOperation operation, ICurrentTenantContext tenants,
  IEntraGraphClientFactory clients, ILogger<ClaimsMappingPolicyService> logger) : IClaimsMappingPolicyService
{
  private const string Collection = "policies/claimsMappingPolicies";
  private static readonly string[] PolicyFields = ["id", "displayName", "definition", "isOrganizationDefault"];

  public Result<ClaimsMappingPolicyCreationContext, GraphOperationError> BeginCreate() => tenants.CurrentTenant.HasValue
    ? new ClaimsMappingPolicyCreationContext(tenants.CurrentTenant.Value.TenantId, operation.SelectionId)
    : new GraphOperationError(GraphOperationErrorType.TenantNotSelected, "Select a tenant to create a policy.");

  public async Task<Result<ClaimsMappingPolicyListItem, GraphOperationError>> CreateAsync(
    ClaimsMappingPolicyCreationContext context, CreateClaimsMappingPolicyRequest request, CancellationToken cancellationToken = default)
  {
    cancellationToken.ThrowIfCancellationRequested();
    if (!IsCurrent(context))
      return DraftTenantChanged();
    var validated = request.ValidateAndNormalize();
    if (validated.IsFailure)
      return new GraphOperationError(GraphOperationErrorType.InvalidInput, validated.Error);
    var policy = new GraphPolicy
    {
      DisplayName = validated.Value.DisplayName,
      Definition = [ClaimsMappingPolicyDefinitionSerializer.Serialize(validated.Value)],
      IsOrganizationDefault = false
    };

    try
    {
      var createdClient = await clients.CreateAsync(cancellationToken);
      if (createdClient.IsFailure)
        return !IsCurrent(context) ? DraftTenantChanged() : CreationFailure(createdClient.Error);
      using var client = createdClient.Value;
      if (!IsCurrent(context))
        return DraftTenantChanged();
      cancellationToken.ThrowIfCancellationRequested();
      // The factory binds this client to the starting tenant. Never resolve a new client after this check.
      var created = await client.Policies.ClaimsMappingPolicies.PostAsync(policy, cancellationToken: cancellationToken);
      if (!IsCurrent(context))
        return DraftTenantChanged();
      if (created is null || !HasValidId(created))
        return CreationFailure(new(GraphOperationErrorType.GraphFailure,
          "Microsoft Graph did not return a valid created policy. Refresh the list before trying again."));
      return MapPolicy(created);
    }
    catch (MicrosoftIdentityWebChallengeUserException exception)
    {
      return CreationFailure(GraphErrorMapping.From(exception.MsalUiRequiredException));
    }
    catch (MsalUiRequiredException exception)
    {
      return CreationFailure(GraphErrorMapping.From(exception));
    }
    catch (MsalException exception)
    {
      return CreationFailure(GraphErrorMapping.From(exception));
    }
    catch (ApiException exception)
    {
      var error = GraphErrorMapping.From(exception);
      return CreationFailure(error with
      {
        Type = exception.ResponseStatusCode is 400 or 409 or 422 ? GraphOperationErrorType.InvalidInput : error.Type,
        Message = exception.ResponseStatusCode switch
        {
          400 or 409 or 422 => "Microsoft Graph rejected the policy. Check its name, attributes, and output claim types.",
          404 => "The claims mapping policy creation endpoint is unavailable in this tenant.",
          _ when error.Type == GraphOperationErrorType.GraphFailure => "Microsoft Graph could not confirm policy creation. Refresh the list before trying again.",
          _ => error.Message
        }
      });
    }
    catch (HttpRequestException)
    {
      return CreationFailure(new(GraphOperationErrorType.GraphFailure,
        "Microsoft Graph could not be reached. Creation may have completed; refresh the list before trying again."));
    }
    catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
    {
      return CreationFailure(new(GraphOperationErrorType.GraphFailure,
        "Policy creation timed out. Creation may have completed; refresh the list before trying again."));
    }
  }

  private bool IsCurrent(ClaimsMappingPolicyCreationContext context) => tenants.CurrentTenant.HasValue
    && tenants.CurrentTenant.Value.TenantId == context.TenantId && operation.SelectionId == context.SelectionId;

  private static GraphOperationError DraftTenantChanged() => new(GraphOperationErrorType.TenantChanged,
    "The tenant changed. Start a new policy draft. If creation was already submitted, check the original tenant's policy list.");

  private GraphOperationError CreationFailure(GraphOperationError error)
  {
    if (error.Type is GraphOperationErrorType.Forbidden or GraphOperationErrorType.ConsentRequired)
      error = error with { Message = "Policy creation requires delegated Policy.ReadWrite.ApplicationConfiguration admin consent and appropriate directory privileges in this tenant. Authenticate the tenant after consent is granted." };
    logger.LogWarning("Claims mapping policy creation failed: {ErrorType}, HTTP {Status}, code {Code}, request {RequestId}",
      error.Type, error.HttpStatus, error.Code, error.RequestId);
    return error;
  }

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
