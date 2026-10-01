using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components;

// Owns cancellation and data lifetime for each tenant-bound list or details page.
public sealed class TenantViewState<T> : IDisposable where T : class
{
  private readonly ICurrentTenantContext _tenants;
  private readonly Action _tenantChanged;
  private CancellationTokenSource? _load;
  private bool _disposed;

  public TenantViewState(ICurrentTenantContext tenants, Action tenantChanged)
  {
    _tenants = tenants;
    _tenantChanged = tenantChanged;
    tenants.TenantChanged += OnTenantChanged;
  }

  public T? Value { get; private set; }
  public GraphOperationError? Error { get; private set; }
  public bool IsLoading { get; private set; }
  public bool HasLoaded { get; private set; }

  public async Task LoadAsync(Func<CancellationToken, Task<Result<T, GraphOperationError>>> load)
  {
    if (_disposed)
      return;
    Clear();
    using var cancellation = new CancellationTokenSource();
    _load = cancellation;
    IsLoading = true;
    try
    {
      var result = await load(cancellation.Token);
      // Operation identity also rejects A -> B -> A and out-of-order search responses.
      if (_disposed || cancellation.IsCancellationRequested || _load != cancellation)
        return;
      Value = result.IsSuccess ? result.Value : null;
      Error = result.IsFailure ? result.Error : null;
      HasLoaded = true;
    }
    catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
    {
      // A new request, tenant selection, or disposal superseded this load.
    }
    finally
    {
      if (_load == cancellation)
      {
        _load = null;
        IsLoading = false;
      }
    }
  }

  private void Clear()
  {
    _load?.Cancel();
    _load = null;
    Value = null;
    Error = null;
    IsLoading = false;
    HasLoaded = false;
  }

  private void OnTenantChanged()
  {
    Clear();
    _tenantChanged();
  }

  public void Dispose()
  {
    _disposed = true;
    _tenants.TenantChanged -= OnTenantChanged;
    Clear();
  }
}
