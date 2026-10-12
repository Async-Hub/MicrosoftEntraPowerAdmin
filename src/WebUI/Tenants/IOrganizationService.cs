using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Infrastructure;
using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public interface IOrganizationService
{
  Task<Result<MicrosoftEntraOrganization, GraphOperationError>> GetCurrentAsync(
    CancellationToken cancellationToken = default);
}
