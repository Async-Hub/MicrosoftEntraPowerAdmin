using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

// Shared tenant lifetime and error boundary for the two directory browsing services.
public sealed class DirectoryReadOperation : IDisposable
{
  private readonly ICurrentTenantContext _tenants;
  private readonly IEntraGraphClientFactory _clients;
  private readonly ILogger<DirectoryReadOperation> _logger;
  private Guid _selectionId = Guid.NewGuid();

  public DirectoryReadOperation(ICurrentTenantContext tenants, IEntraGraphClientFactory clients,
    ILogger<DirectoryReadOperation> logger)
  {
    _tenants = tenants;
    _clients = clients;
    _logger = logger;
    tenants.TenantChanged += OnTenantChanged;
  }

  internal async Task<Result<T, GraphOperationError>> RunAsync<T>(
    Func<GraphServiceClient, Task<Result<T, GraphOperationError>>> read,
    CancellationToken cancellationToken)
  {
    cancellationToken.ThrowIfCancellationRequested();
    var tenant = _tenants.CurrentTenant;
    var selection = _selectionId;
    if (tenant.HasNoValue)
      return new GraphOperationError(GraphOperationErrorType.TenantNotSelected, "Select a tenant to browse applications.");

    try
    {
      var created = await _clients.CreateAsync(cancellationToken);
      if (created.IsFailure)
        return created.Error;
      using var client = created.Value;
      if (selection != _selectionId)
        return TenantChanged();
      var result = await read(client);
      cancellationToken.ThrowIfCancellationRequested();
      return selection == _selectionId ? result : TenantChanged();
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
          403 => "Access denied. Ask an administrator to grant delegated Application.Read.All admin consent in this tenant and check your directory privileges.",
          404 => "The object was not found in the current tenant.",
          _ when error.Type == GraphOperationErrorType.GraphFailure => "Microsoft Graph could not retrieve the applications. Try again later.",
          _ => error.Message
        }
      });
    }
    catch (HttpRequestException)
    {
      return Failure(new(GraphOperationErrorType.GraphFailure, "Microsoft Graph could not be reached. Try again later."));
    }
    catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
    {
      return Failure(new(GraphOperationErrorType.GraphFailure, "Microsoft Graph timed out. Try again later."));
    }
  }

  internal bool IsValid(DirectoryContinuation continuation, string collection) =>
    continuation.SelectionId == _selectionId && IsCollectionUrl(continuation.Url, collection);

  internal Result<DirectoryContinuation?, GraphOperationError> Continuation(string? url, string collection)
  {
    if (string.IsNullOrEmpty(url))
      return Result.Success<DirectoryContinuation?, GraphOperationError>(null);
    if (!IsCollectionUrl(url, collection))
      return new GraphOperationError(GraphOperationErrorType.GraphFailure, "Microsoft Graph returned an invalid next page.");
    return new DirectoryContinuation(url, _selectionId);
  }

  private static bool IsCollectionUrl(string url, string collection) =>
    Uri.TryCreate(url, UriKind.Absolute, out var uri) && uri.Scheme == "https"
    && uri.Host == "graph.microsoft.com" && uri.IsDefaultPort && uri.UserInfo.Length == 0
    && uri.AbsolutePath == $"/v1.0/{collection}" && uri.Fragment.Length == 0;

  internal static GraphOperationError InvalidId() => new(GraphOperationErrorType.InvalidInput, "Enter a non-empty GUID.");
  internal static GraphOperationError TenantChanged() => new(GraphOperationErrorType.TenantChanged, "The tenant changed. Start a new search.");
  private void OnTenantChanged() => _selectionId = Guid.NewGuid();

  private GraphOperationError Failure(GraphOperationError error)
  {
    _logger.LogWarning("Directory read failed: {ErrorType}, HTTP {Status}, code {Code}, request {RequestId}",
      error.Type, error.HttpStatus, error.Code, error.RequestId);
    return error;
  }

  public void Dispose() => _tenants.TenantChanged -= OnTenantChanged;
}
