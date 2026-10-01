using CSharpFunctionalExtensions;
using GraphModel = Microsoft.Graph.Models.Application;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed class ApplicationService(DirectoryReadOperation operation) : IApplicationService
{
  private static readonly string[] ListFields = ["id", "appId", "displayName", "signInAudience", "createdDateTime", "publisherDomain"];
  private static readonly string[] DetailFields = [.. ListFields, "description", "identifierUris", "web", "spa", "publicClient", "tags", "api"];

  public Task<Result<DirectoryPage<ApplicationListItem>, GraphOperationError>> SearchAsync(
    string? search, DirectoryContinuation? nextPage = null, CancellationToken cancellationToken = default) =>
    operation.RunAsync<DirectoryPage<ApplicationListItem>>(async client =>
    {
      if (nextPage is not null && !operation.IsValid(nextPage, "applications"))
        return DirectoryReadOperation.TenantChanged();
      var term = search?.Trim() ?? "";
      if (term.Length > 200)
        return new GraphOperationError(GraphOperationErrorType.InvalidInput, "Search text must be 200 characters or fewer.");
      var filter = Guid.TryParse(term, out var id)
        ? $"appId eq '{id:D}' or id eq '{id:D}'"
        : term.Length == 0 ? null : $"startswith(displayName,'{term.Replace("'", "''")}')";
      var response = nextPage is null
        ? await client.Applications.GetAsync(request =>
        {
          request.QueryParameters.Select = ListFields;
          request.QueryParameters.Filter = filter;
          request.QueryParameters.Top = 25;
          request.QueryParameters.Count = true;
          request.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken)
        : await client.Applications.WithUrl(nextPage.Url).GetAsync(request =>
          request.Headers.Add("ConsistencyLevel", "eventual"), cancellationToken);

      if (response?.Value is not { } items || items.Any(item => !HasValidObjectId(item)))
        return InvalidResponse();
      var continuation = operation.Continuation(response.OdataNextLink, "applications");
      if (continuation.IsFailure)
        return continuation.Error;
      return new DirectoryPage<ApplicationListItem>(
        Array.AsReadOnly(items.Select(MapListItem).ToArray()), continuation.Value);
    }, cancellationToken);

  public Task<Result<ApplicationDetails, GraphOperationError>> GetByObjectIdAsync(
    Guid objectId, CancellationToken cancellationToken = default) =>
    operation.RunAsync<ApplicationDetails>(async client =>
    {
      if (objectId == Guid.Empty)
        return DirectoryReadOperation.InvalidId();
      var model = await client.Applications[objectId.ToString("D")].GetAsync(request =>
        request.QueryParameters.Select = DetailFields, cancellationToken);
      if (model is null)
        return new GraphOperationError(GraphOperationErrorType.NotFound, "The object was not found in the current tenant.");
      if (!HasValidObjectId(model) || Guid.Parse(model.Id!) != objectId)
        return InvalidResponse();
      return MapDetails(model);
    }, cancellationToken);

  public Task<Result<Maybe<ApplicationDetails>, GraphOperationError>> FindByAppIdAsync(
    Guid appId, CancellationToken cancellationToken = default) =>
    operation.RunAsync<Maybe<ApplicationDetails>>(async client =>
    {
      if (appId == Guid.Empty)
        return DirectoryReadOperation.InvalidId();
      var response = await client.Applications.GetAsync(request =>
      {
        request.QueryParameters.Select = DetailFields;
        request.QueryParameters.Filter = $"appId eq '{appId:D}'";
        request.QueryParameters.Top = 2;
      }, cancellationToken);
      if (response?.Value is not { } items || items.Count > 1 || !string.IsNullOrEmpty(response.OdataNextLink))
        return InvalidResponse();
      if (items.Count == 0)
        return Maybe<ApplicationDetails>.None;
      if (!HasValidObjectId(items[0]) || !Guid.TryParse(items[0].AppId, out var actualAppId) || actualAppId != appId)
        return InvalidResponse();
      return Maybe<ApplicationDetails>.From(MapDetails(items[0]));
    }, cancellationToken);

  private static bool HasValidObjectId(GraphModel model) => Guid.TryParse(model.Id, out var id) && id != Guid.Empty;
  private static GraphOperationError InvalidResponse() => new(GraphOperationErrorType.GraphFailure,
    "Microsoft Graph returned invalid or ambiguous application data.");

  // All callers validate the object ID before mapping. Client ID may be absent.
  private static ApplicationListItem MapListItem(GraphModel model) => new(
    Guid.Parse(model.Id!), Guid.TryParse(model.AppId, out var appId) && appId != Guid.Empty ? appId : null,
    string.IsNullOrWhiteSpace(model.DisplayName) ? "Name unavailable" : model.DisplayName,
    model.SignInAudience, model.CreatedDateTime, model.PublisherDomain);

  private static ApplicationDetails MapDetails(GraphModel model) => new(
    MapListItem(model), model.Description,
    Array.AsReadOnly(model.IdentifierUris?.ToArray() ?? []),
    Array.AsReadOnly(model.Web?.RedirectUris?.ToArray() ?? []),
    Array.AsReadOnly(model.Spa?.RedirectUris?.ToArray() ?? []),
    Array.AsReadOnly(model.PublicClient?.RedirectUris?.ToArray() ?? []),
    Array.AsReadOnly(model.Tags?.ToArray() ?? []),
    model.Api?.AcceptMappedClaims, model.Api?.RequestedAccessTokenVersion);
}
