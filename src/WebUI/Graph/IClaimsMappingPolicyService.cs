using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public interface IClaimsMappingPolicyService
{
  Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> ListAsync(CancellationToken cancellationToken = default);
  Task<Result<ClaimsMappingPolicyDetails, GraphOperationError>> GetByObjectIdAsync(Guid objectId, CancellationToken cancellationToken = default);
  Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> FindForServicePrincipalAsync(
    Guid servicePrincipalObjectId, CancellationToken cancellationToken = default);
}
