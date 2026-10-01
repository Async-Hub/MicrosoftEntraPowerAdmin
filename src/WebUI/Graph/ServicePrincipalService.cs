using CSharpFunctionalExtensions;
using GraphModel = Microsoft.Graph.Models.ServicePrincipal;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed class ServicePrincipalService(DirectoryReadOperation operation) : IServicePrincipalService
{
  private static readonly string[] ListFields = ["id", "appId", "displayName", "servicePrincipalType", "accountEnabled", "verifiedPublisher"];
  private static readonly string[] DetailFields = [.. ListFields, "description", "homepage", "loginUrl", "servicePrincipalNames", "appOwnerOrganizationId", "tags"];

  public Task<Result<DirectoryPage<ServicePrincipalListItem>, GraphOperationError>> SearchAsync(
    string? search, DirectoryContinuation? nextPage = null, CancellationToken cancellationToken = default) =>
    operation.RunAsync<DirectoryPage<ServicePrincipalListItem>>(async client =>
    {
      if (nextPage is not null && !operation.IsValid(nextPage, "servicePrincipals"))
        return DirectoryReadOperation.TenantChanged();
      var term = search?.Trim() ?? "";
      if (term.Length > 200)
        return new GraphOperationError(GraphOperationErrorType.InvalidInput, "Search text must be 200 characters or fewer.");
      var filter = Guid.TryParse(term, out var id)
        ? $"appId eq '{id:D}' or id eq '{id:D}'"
        : term.Length == 0 ? null : $"startswith(displayName,'{term.Replace("'", "''")}')";
      var response = nextPage is null
        ? await client.ServicePrincipals.GetAsync(request =>
        {
          request.QueryParameters.Select = ListFields;
          request.QueryParameters.Filter = filter;
          request.QueryParameters.Top = 25;
          request.QueryParameters.Count = true;
          request.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken)
        : await client.ServicePrincipals.WithUrl(nextPage.Url).GetAsync(request =>
          request.Headers.Add("ConsistencyLevel", "eventual"), cancellationToken);

      if (response?.Value is not { } items || items.Any(item => !HasValidObjectId(item)))
        return InvalidResponse();
      var continuation = operation.Continuation(response.OdataNextLink, "servicePrincipals");
      if (continuation.IsFailure)
        return continuation.Error;
      return new DirectoryPage<ServicePrincipalListItem>(
        Array.AsReadOnly(items.Select(MapListItem).ToArray()), continuation.Value);
    }, cancellationToken);

  public Task<Result<ServicePrincipalDetails, GraphOperationError>> GetByObjectIdAsync(
    Guid objectId, CancellationToken cancellationToken = default) =>
    operation.RunAsync<ServicePrincipalDetails>(async client =>
    {
      if (objectId == Guid.Empty)
        return DirectoryReadOperation.InvalidId();
      var model = await client.ServicePrincipals[objectId.ToString("D")].GetAsync(request =>
        request.QueryParameters.Select = DetailFields, cancellationToken);
      if (model is null)
        return new GraphOperationError(GraphOperationErrorType.NotFound, "The object was not found in the current tenant.");
      if (!HasValidObjectId(model) || Guid.Parse(model.Id!) != objectId)
        return InvalidResponse();
      return MapDetails(model);
    }, cancellationToken);

  public Task<Result<Maybe<ServicePrincipalDetails>, GraphOperationError>> FindByAppIdAsync(
    Guid appId, CancellationToken cancellationToken = default) =>
    operation.RunAsync<Maybe<ServicePrincipalDetails>>(async client =>
    {
      if (appId == Guid.Empty)
        return DirectoryReadOperation.InvalidId();
      var response = await client.ServicePrincipals.GetAsync(request =>
      {
        request.QueryParameters.Select = DetailFields;
        request.QueryParameters.Filter = $"appId eq '{appId:D}'";
        request.QueryParameters.Top = 2;
      }, cancellationToken);
      if (response?.Value is not { } items || items.Count > 1 || !string.IsNullOrEmpty(response.OdataNextLink))
        return InvalidResponse();
      if (items.Count == 0)
        return Maybe<ServicePrincipalDetails>.None;
      if (!HasValidObjectId(items[0]) || !Guid.TryParse(items[0].AppId, out var actualAppId) || actualAppId != appId)
        return InvalidResponse();
      return Maybe<ServicePrincipalDetails>.From(MapDetails(items[0]));
    }, cancellationToken);

  private static bool HasValidObjectId(GraphModel model) => Guid.TryParse(model.Id, out var id) && id != Guid.Empty;
  private static GraphOperationError InvalidResponse() => new(GraphOperationErrorType.GraphFailure,
    "Microsoft Graph returned invalid or ambiguous service principal data.");

  // All callers validate the object ID before mapping. Client ID may be absent.
  internal static ServicePrincipalListItem MapListItem(GraphModel model) => new(
    Guid.Parse(model.Id!), Guid.TryParse(model.AppId, out var appId) && appId != Guid.Empty ? appId : null,
    string.IsNullOrWhiteSpace(model.DisplayName) ? "Name unavailable" : model.DisplayName,
    model.ServicePrincipalType, model.AccountEnabled, model.VerifiedPublisher?.DisplayName);

  private static ServicePrincipalDetails MapDetails(GraphModel model) => new(
    MapListItem(model), model.Description,
    model.Homepage, model.LoginUrl,
    Array.AsReadOnly(model.ServicePrincipalNames?.ToArray() ?? []),
    model.AppOwnerOrganizationId, Array.AsReadOnly(model.Tags?.ToArray() ?? []));
}
