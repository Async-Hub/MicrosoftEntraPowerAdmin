using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public interface IServicePrincipalService
{
  Task<Result<DirectoryPage<ServicePrincipalListItem>, GraphOperationError>> SearchAsync(
    string? search, DirectoryContinuation? nextPage = null, CancellationToken cancellationToken = default);
  Task<Result<ServicePrincipalDetails, GraphOperationError>> GetByObjectIdAsync(
    Guid objectId, CancellationToken cancellationToken = default);
  Task<Result<Maybe<ServicePrincipalDetails>, GraphOperationError>> FindByAppIdAsync(
    Guid appId, CancellationToken cancellationToken = default);
}
