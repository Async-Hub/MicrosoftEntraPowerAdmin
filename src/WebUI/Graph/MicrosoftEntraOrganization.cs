namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed record MicrosoftEntraOrganization(
  Guid TenantId,
  string DisplayName,
  IReadOnlyList<string> VerifiedDomains);
