using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class DeleteClaimsMappingPolicyDialog(ICurrentTenantContext tenants, IClaimsMappingPolicyService service) : IDisposable
{
  [Parameter, EditorRequired] public ClaimsMappingPolicyDeletionContext Context { get; set; } = default!;
  [Parameter] public EventCallback<GraphOperationError?> OnFinished { get; set; }
  [Parameter] public EventCallback OnCancel { get; set; }

  private readonly CancellationTokenSource _lifetime = new();
  private string _tenantName = "";
  private string _confirmation = "";
  private bool _isDeleting;
  private bool _invalidated;
  private bool _disposed;
  private GraphOperationError? _error;
  private bool CanDelete => !_isDeleting && !_invalidated && !_disposed && Context.CanConfirm(_confirmation);
  private string IncludeBasicClaimSet => Context.Original.Policy.Definitions is [var definition]
    && definition.Document?.IncludeBasicClaimSet is { } include ? include ? "Yes" : "No" : "Unavailable";

  protected override void OnInitialized()
  {
    tenants.TenantChanged += OnTenantChanged;
    if (tenants.CurrentTenant.HasNoValue || tenants.CurrentTenant.Value.TenantId != Context.TenantId)
    {
      OnTenantChanged();
      return;
    }
    _tenantName = tenants.CurrentTenant.Value.Name;
  }

  private Task CancelAsync() => _isDeleting ? Task.CompletedTask : OnCancel.InvokeAsync();

  private async Task DeleteAsync()
  {
    if (!CanDelete)
      return;
    _isDeleting = true;
    try
    {
      var result = await service.DeleteAsync(Context, _confirmation, _lifetime.Token);
      if (_invalidated || _disposed)
        return;
      // Every attempt closes this confirmation. The parent reloads Graph state before another attempt.
      _invalidated = true;
      _confirmation = "";
      await OnFinished.InvokeAsync(result.IsFailure ? result.Error : null);
    }
    catch (OperationCanceledException) when (_lifetime.IsCancellationRequested)
    {
      // Tenant switching or navigation discarded this tenant-bound confirmation.
    }
    finally
    {
      _isDeleting = false;
    }
  }

  private void OnTenantChanged()
  {
    if (_disposed)
      return;
    _invalidated = true;
    _confirmation = "";
    _lifetime.Cancel();
    _error = new(GraphOperationErrorType.TenantChanged, "The tenant changed. Reopen the policy in the selected tenant.");
    _ = InvokeAsync(StateHasChanged);
  }

  public void Dispose()
  {
    _disposed = true;
    tenants.TenantChanged -= OnTenantChanged;
    _lifetime.Cancel();
    _lifetime.Dispose();
  }
}
