using CSharpFunctionalExtensions;
using Microsoft.Graph;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public interface IEntraGraphClientFactory
{
  Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(CancellationToken cancellationToken = default);
}
