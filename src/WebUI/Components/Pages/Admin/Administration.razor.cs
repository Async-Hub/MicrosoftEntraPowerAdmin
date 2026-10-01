using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class Administration(
  ICurrentTenantContext tenantContext,
  IOrganizationService organizationService) : IDisposable
{
  private TenantConnectionInfo _connection = new(TenantConnectionState.NoTenantSelected);
  private CancellationTokenSource? _loadCancellation;
  private bool _disposed;
  private bool _restoringSelection;
  private Guid? _appliedReturnTenant;

  [SupplyParameterFromQuery(Name = "tenant")]
  public Guid? ReturnTenantId { get; set; }

  private string StatusText => _connection.State switch
  {
    TenantConnectionState.NoTenantSelected => "No tenant selected",
    TenantConnectionState.Connecting => "Connecting",
    TenantConnectionState.Connected => "Connected",
    TenantConnectionState.InteractionRequired => "Interaction required",
    TenantConnectionState.AccessDenied => "Access denied",
    _ => "Graph error"
  };

  protected override void OnInitialized() => tenantContext.TenantChanged += OnTenantChanged;

  protected override async Task OnParametersSetAsync()
  {
    if (ReturnTenantId.HasValue && ReturnTenantId != _appliedReturnTenant)
    {
      _appliedReturnTenant = ReturnTenantId;
      _loadCancellation?.Cancel();
      // The query only restores navigation after authentication; the allowlist is authoritative.
      _restoringSelection = true;
      var selected = tenantContext.SelectTenant(ReturnTenantId.Value);
      _restoringSelection = false;
      if (selected.IsFailure)
      {
        _connection = new(TenantConnectionState.GraphError, Error:
          new(GraphOperationErrorType.TenantInaccessible, selected.Error));
        return;
      }
    }
    await LoadAsync();
  }

  private void OnTenantChanged()
  {
    if (_restoringSelection || _disposed)
      return;

    _ = InvokeAsync(async () =>
    {
      try
      {
        await LoadAsync();
      }
      catch (Exception exception)
      {
        // Forward unexpected failures from the event task to Blazor's error boundary.
        await DispatchExceptionAsync(exception);
      }
    });
  }

  private async Task LoadAsync()
  {
    if (_disposed)
      return;

    _loadCancellation?.Cancel();
    var tenant = tenantContext.CurrentTenant;
    _connection = tenant.HasValue
      ? new(TenantConnectionState.Connecting, tenant.Value)
      : new(TenantConnectionState.NoTenantSelected);
    StateHasChanged();
    if (tenant.HasNoValue)
      return;

    using var cancellation = new CancellationTokenSource();
    _loadCancellation = cancellation;
    try
    {
      var result = await organizationService.GetCurrentAsync(cancellation.Token);
      // The operation identity also protects A -> B -> A switches and out-of-order retries.
      if (_disposed || cancellation.IsCancellationRequested || _loadCancellation != cancellation
        || tenantContext.CurrentTenant != tenant)
        return;

      _connection = result.IsSuccess
        ? new(TenantConnectionState.Connected, tenant.Value, result.Value)
        : new(StateFor(result.Error), tenant.Value, Error: result.Error);
      StateHasChanged();
    }
    catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
    {
      // A new selection or component disposal superseded this operation.
    }
    finally
    {
      if (_loadCancellation == cancellation)
        _loadCancellation = null;
    }
  }

  private static TenantConnectionState StateFor(GraphOperationError error) => error.Type switch
  {
    GraphOperationErrorType.InteractionRequired or GraphOperationErrorType.ConsentRequired
      or GraphOperationErrorType.Unauthorized => TenantConnectionState.InteractionRequired,
    GraphOperationErrorType.Forbidden or GraphOperationErrorType.TenantInaccessible => TenantConnectionState.AccessDenied,
    _ => TenantConnectionState.GraphError
  };

  public void Dispose()
  {
    _disposed = true;
    tenantContext.TenantChanged -= OnTenantChanged;
    _loadCancellation?.Cancel();
  }
}
