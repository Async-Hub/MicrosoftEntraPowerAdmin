using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class Applications(ICurrentTenantContext tenants, IApplicationService service,
  NavigationManager navigation) : IDisposable
{
  private TenantViewState<DirectoryPage<ApplicationListItem>> _state = default!;
  private string? _search;
  private bool HasTenant => tenants.CurrentTenant.HasValue;

  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);

  private Task SearchAsync() => _state.LoadAsync(token => service.SearchAsync(_search, cancellationToken: token));
  private Task NextAsync(DirectoryContinuation next) =>
    _state.LoadAsync(token => service.SearchAsync(_search, next, token));

  private void Open(Guid objectId)
  {
    // Only objects still present in this page's current tenant state can be opened.
    if (_state.Value?.Items.Any(item => item.ObjectId == objectId) == true)
      navigation.NavigateTo($"/applications/{objectId:D}");
  }

  private void OnTenantChanged()
  {
    _search = null;
    _ = InvokeAsync(StateHasChanged);
  }

  public void Dispose() => _state.Dispose();
}
