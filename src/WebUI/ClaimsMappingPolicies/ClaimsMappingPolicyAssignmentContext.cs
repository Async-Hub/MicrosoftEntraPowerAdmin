namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.ClaimsMappingPolicies;

// Selection identity invalidates confirmations even after switching A -> B -> A.
public sealed record ClaimsMappingPolicyAssignmentContext(Guid TenantId, Guid SelectionId);
