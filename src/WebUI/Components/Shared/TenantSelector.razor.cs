using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class TenantSelector(ICurrentTenantContext tenantContext) : IDisposable
{
    private ICurrentTenantContext TenantContext => tenantContext;
    private string? _error;
    private Guid? SelectedTenantId => TenantContext.CurrentTenant.HasValue
        ? TenantContext.CurrentTenant.Value.TenantId : null;

    protected override void OnInitialized() => TenantContext.TenantChanged += OnTenantChanged;

    private void SelectTenant(Guid? tenantId)
    {
        var result = TenantContext.SelectTenant(tenantId ?? Guid.Empty);
        _error = result.IsFailure ? result.Error : null;
    }

    private void OnTenantChanged() => _ = InvokeAsync(StateHasChanged);

    public void Dispose() => TenantContext.TenantChanged -= OnTenantChanged;
}
