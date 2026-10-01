using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class ClaimsMappingPolicies(ICurrentTenantContext tenants, IClaimsMappingPolicyService service,
  NavigationManager navigation) : IDisposable
{
  private TenantViewState<IReadOnlyList<ClaimsMappingPolicyListItem>> _state = default!;
  private bool HasTenant => tenants.CurrentTenant.HasValue;
  private ClaimsMappingPolicyCreationContext? _createContext;
  private GraphOperationError? _createError;
  private ClaimsMappingPolicyListItem? _created;
  private bool _draftDiscarded;
  private IReadOnlyList<ClaimsMappingPolicyListItem>? DisplayPolicies => _created is { } created
    ? new[] { created }.Concat(_state.Value?.Where(policy => policy.ObjectId != created.ObjectId) ?? []).ToArray()
    : _state.Value;
  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);
  protected override Task OnInitializedAsync() => LoadAsync();
  private Task LoadAsync() => _state.LoadAsync(service.ListAsync);

  private void StartCreate()
  {
    if (_createContext is not null)
      return;
    var context = service.BeginCreate();
    _createContext = context.IsSuccess ? context.Value : null;
    _createError = context.IsFailure ? context.Error : null;
    _created = null;
    _draftDiscarded = false;
  }

  private void CancelCreate()
  {
    _createContext = null;
    _createError = null;
  }

  private async Task PolicyCreatedAsync(ClaimsMappingPolicyListItem policy)
  {
    if (_createContext is null)
      return;
    _createContext = null;
    _created = policy;
    // Keep the POST response authoritative, including during eventual consistency or a failed refresh.
    await LoadAsync();
  }

  private void Open(Guid objectId)
  {
    if (DisplayPolicies?.Any(item => item.ObjectId == objectId) == true)
      navigation.NavigateTo($"/claims-mapping-policies/{objectId:D}");
  }

  private void OnTenantChanged()
  {
    _draftDiscarded = _createContext is not null;
    _createContext = null;
    _createError = null;
    _created = null;
    _ = InvokeAsync(async () =>
    {
      var load = LoadAsync();
      StateHasChanged();
      await load;
      StateHasChanged();
    });
  }

  public void Dispose() => _state.Dispose();
}
