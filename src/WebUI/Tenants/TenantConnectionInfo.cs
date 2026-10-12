using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Infrastructure;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public sealed record TenantConnectionInfo(
  TenantConnectionState State,
  MicrosoftEntraTenant? Tenant = null,
  MicrosoftEntraOrganization? Organization = null,
  GraphOperationError? Error = null);