using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public interface IOrganizationService
{
  Task<Result<MicrosoftEntraOrganization, GraphOperationError>> GetCurrentAsync(
    CancellationToken cancellationToken = default);
}
