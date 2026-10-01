using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class AssignedClaimsMappingPolicies(ICurrentTenantContext tenants, IClaimsMappingPolicyService service,
  NavigationManager navigation) : IDisposable
{
  [Parameter, EditorRequired] public Guid ServicePrincipalObjectId { get; set; }
  private TenantViewState<IReadOnlyList<ClaimsMappingPolicyListItem>> _state = default!;
  private bool _tenantChanged;
  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);
  protected override Task OnParametersSetAsync() => LoadAsync();
  private Task LoadAsync() => _tenantChanged ? Task.CompletedTask :
    _state.LoadAsync(token => service.FindForServicePrincipalAsync(ServicePrincipalObjectId, token));

  private void Open(Guid objectId)
  {
    if (_state.Value?.Any(item => item.ObjectId == objectId) == true)
      navigation.NavigateTo($"/claims-mapping-policies/{objectId:D}");
  }

  private void OnTenantChanged()
  {
    _tenantChanged = true;
    _ = InvokeAsync(StateHasChanged);
  }

  public void Dispose() => _state.Dispose();
}
