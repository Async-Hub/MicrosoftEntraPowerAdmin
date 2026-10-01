using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public interface IApplicationService
{
  Task<Result<DirectoryPage<ApplicationListItem>, GraphOperationError>> SearchAsync(
    string? search, DirectoryContinuation? nextPage = null, CancellationToken cancellationToken = default);
  Task<Result<ApplicationDetails, GraphOperationError>> GetByObjectIdAsync(
    Guid objectId, CancellationToken cancellationToken = default);
  Task<Result<Maybe<ApplicationDetails>, GraphOperationError>> FindByAppIdAsync(
    Guid appId, CancellationToken cancellationToken = default);
}
