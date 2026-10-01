using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public interface IClaimsMappingPolicyService
{
  Result<ClaimsMappingPolicyAssignmentContext, GraphOperationError> BeginAssignment();
  Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> GetAssignedPoliciesAsync(
    ClaimsMappingPolicyAssignmentContext context, Guid servicePrincipalObjectId, CancellationToken cancellationToken = default);
  Task<UnitResult<GraphOperationError>> AssignAsync(ClaimsMappingPolicyAssignmentContext context,
    Guid servicePrincipalObjectId, Guid claimsMappingPolicyObjectId, CancellationToken cancellationToken = default);
  Task<UnitResult<GraphOperationError>> UnassignAsync(ClaimsMappingPolicyAssignmentContext context,
    Guid servicePrincipalObjectId, Guid claimsMappingPolicyObjectId, CancellationToken cancellationToken = default);
  Result<ClaimsMappingPolicyCreationContext, GraphOperationError> BeginCreate();
  Task<Result<ClaimsMappingPolicyListItem, GraphOperationError>> CreateAsync(
    ClaimsMappingPolicyCreationContext context, CreateClaimsMappingPolicyRequest request, CancellationToken cancellationToken = default);
  Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> ListAsync(CancellationToken cancellationToken = default);
  Task<Result<ClaimsMappingPolicyDetails, GraphOperationError>> GetByObjectIdAsync(Guid objectId, CancellationToken cancellationToken = default);
  Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>> FindForServicePrincipalAsync(
    Guid servicePrincipalObjectId, CancellationToken cancellationToken = default);
}
