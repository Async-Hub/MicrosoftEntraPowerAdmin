using CSharpFunctionalExtensions;
using Microsoft.Graph;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Infrastructure;

public interface IEntraGraphClientFactory
{
  Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(CancellationToken cancellationToken = default);
}
