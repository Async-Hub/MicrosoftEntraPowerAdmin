namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

// Selection identity invalidates confirmations even after switching A -> B -> A.
public sealed record ClaimsMappingPolicyAssignmentContext(Guid TenantId, Guid SelectionId);
