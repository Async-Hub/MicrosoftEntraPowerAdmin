using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class DirectoryStatus(ICurrentTenantContext tenants) : IDisposable
{
  [Parameter] public GraphOperationError? Error { get; set; }
  [Parameter] public bool IsLoading { get; set; }
  private string TenantName => tenants.CurrentTenant.HasValue ? tenants.CurrentTenant.Value.Name : "None selected";
  private Guid? TenantId => tenants.CurrentTenant.HasValue ? tenants.CurrentTenant.Value.TenantId : null;
  private bool NeedsAuthentication => Error?.Type is GraphOperationErrorType.InteractionRequired
    or GraphOperationErrorType.ConsentRequired or GraphOperationErrorType.Unauthorized;

  protected override void OnInitialized() => tenants.TenantChanged += OnTenantChanged;
  private void OnTenantChanged() => _ = InvokeAsync(StateHasChanged);
  public void Dispose() => tenants.TenantChanged -= OnTenantChanged;
}
