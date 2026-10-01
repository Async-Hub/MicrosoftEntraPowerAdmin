namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public enum TenantConnectionState
{
  NoTenantSelected,
  Connecting,
  Connected,
  InteractionRequired,
  AccessDenied,
  GraphError
}