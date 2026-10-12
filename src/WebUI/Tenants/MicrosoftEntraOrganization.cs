namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public sealed record MicrosoftEntraOrganization(
  Guid TenantId,
  string DisplayName,
  IReadOnlyList<string> VerifiedDomains);
