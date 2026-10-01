using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class ClaimsMappingPolicies(ICurrentTenantContext tenants, IClaimsMappingPolicyService service,
  NavigationManager navigation) : IDisposable
{
  private TenantViewState<IReadOnlyList<ClaimsMappingPolicyListItem>> _state = default!;
  private bool HasTenant => tenants.CurrentTenant.HasValue;
  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);
  protected override Task OnInitializedAsync() => LoadAsync();
  private Task LoadAsync() => _state.LoadAsync(service.ListAsync);

  private void Open(Guid objectId)
  {
    if (_state.Value?.Any(item => item.ObjectId == objectId) == true)
      navigation.NavigateTo($"/claims-mapping-policies/{objectId:D}");
  }

  private void OnTenantChanged() => _ = InvokeAsync(async () =>
  {
    var load = LoadAsync();
    StateHasChanged();
    await load;
    StateHasChanged();
  });

  public void Dispose() => _state.Dispose();
}
