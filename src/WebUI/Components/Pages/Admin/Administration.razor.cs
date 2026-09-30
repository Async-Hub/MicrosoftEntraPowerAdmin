using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class Administration(ICurrentTenantContext tenantContext) : IDisposable
{
    private ICurrentTenantContext TenantContext => tenantContext;

    protected override void OnInitialized() => TenantContext.TenantChanged += OnTenantChanged;

    private void OnTenantChanged() => _ = InvokeAsync(StateHasChanged);

    public void Dispose() => TenantContext.TenantChanged -= OnTenantChanged;
}
