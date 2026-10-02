using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class CreateClaimsMappingPolicy(ICurrentTenantContext tenants, IClaimsMappingPolicyService service) : IDisposable
{
  [Parameter, EditorRequired] public ClaimsMappingPolicyCreationContext Context { get; set; } = default!;
  [Parameter] public EventCallback<ClaimsMappingPolicyListItem> OnCreated { get; set; }
  [Parameter] public EventCallback OnCancel { get; set; }

  private readonly ClaimsMappingPolicyDraft _draft = new();
  private string _tenantName = "";
  private string? _preview;
  private CreateClaimsMappingPolicyRequest? _review;
  private GraphOperationError? _error;
  private bool _isCreating;
  private bool _invalidated;
  private bool _disposed;
  private CancellationTokenSource? _creation;
  private bool IsDisabled => _isCreating || _invalidated;

  protected override void OnInitialized()
  {
    _tenantName = tenants.CurrentTenant.HasValue ? tenants.CurrentTenant.Value.Name : "No tenant selected";
    tenants.TenantChanged += OnTenantChanged;
  }

  private void DraftChanged()
  {
    _review = null;
    _preview = null;
    _error = null;
  }

  private void Review()
  {
    if (IsDisabled)
      return;
    var request = _draft.ToRequest().ValidateAndNormalize();
    _error = request.IsFailure ? new(GraphOperationErrorType.InvalidInput, request.Error) : null;
    _review = request.IsSuccess ? request.Value : null;
    _preview = _review is null ? null : ClaimsMappingPolicyDefinitionSerializer.Serialize(_review, indented: true);
  }

  private async Task CreateAsync()
  {
    if (IsDisabled || _review is null)
      return;
    _isCreating = true;
    _error = null;
    using var cancellation = new CancellationTokenSource();
    _creation = cancellation;
    try
    {
      var result = await service.CreateAsync(Context, _review, cancellation.Token);
      if (_disposed || _invalidated)
        return;
      if (result.IsFailure)
      {
        _error = result.Error;
        _invalidated = result.Error.Type == GraphOperationErrorType.TenantChanged;
        return;
      }
      // A successful POST must never be offered for resubmission, even while the list reloads.
      _invalidated = true;
      await OnCreated.InvokeAsync(result.Value);
    }
    catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
    {
      // Tenant selection or navigation superseded this draft; ignore a late response.
    }
    finally
    {
      _creation = null;
      _isCreating = false;
    }
  }

  private void OnTenantChanged()
  {
    _invalidated = true;
    _creation?.Cancel();
    _review = null;
    _preview = null;
    _draft.Claims.Clear();
    _draft.DisplayName = "";
    _error = new(GraphOperationErrorType.TenantChanged, "The tenant changed. Cancel and start a new policy draft.");
    _ = InvokeAsync(StateHasChanged);
  }

  public void Dispose()
  {
    _disposed = true;
    tenants.TenantChanged -= OnTenantChanged;
    _creation?.Cancel();
  }
}
