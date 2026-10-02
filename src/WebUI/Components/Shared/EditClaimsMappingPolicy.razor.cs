using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class EditClaimsMappingPolicy(ICurrentTenantContext tenants, IClaimsMappingPolicyService service) : IDisposable
{
  [Parameter, EditorRequired] public ClaimsMappingPolicyEditContext Context { get; set; } = default!;
  [Parameter] public EventCallback<ClaimsMappingPolicyDetails> OnSaved { get; set; }
  [Parameter] public EventCallback<GraphOperationError> OnMissing { get; set; }
  [Parameter] public EventCallback OnReload { get; set; }
  [Parameter] public EventCallback OnCancel { get; set; }

  private ClaimsMappingPolicyDraft _draft = new();
  private ClaimsMappingDefinitionEditing _editing = default!;
  private string _tenantName = "";
  private UpdateClaimsMappingPolicyRequest? _review;
  private string? _preview;
  private GraphOperationError? _error;
  private bool _definitionChanged;
  private bool _confirmed;
  private bool _isSaving;
  private bool _invalidated;
  private bool _requiresReload;
  private bool _disposed;
  private CancellationTokenSource? _save;
  private bool IsDisabled => _isSaving || _invalidated || _requiresReload;
  private ClaimsMappingAssignments? Assignments => Context.Original.Assignments.IsSuccess ? Context.Original.Assignments.Value : null;
  private IReadOnlyList<Guid> AssignmentIds => Assignments is { } assignments
    ? assignments.ServicePrincipals.Select(item => item.ObjectId).Concat(assignments.Applications.Select(item => item.ObjectId))
      .Concat(assignments.OtherObjects.Select(item => item.ObjectId)).Order().ToArray() : [];
  private bool NeedsConfirmation => _definitionChanged && AssignmentIds.Count > 0;

  protected override void OnInitialized()
  {
    _editing = ClaimsMappingDefinitionEditing.Inspect(Context.Original.Policy);
    _draft = _editing.Definition.HasValue ? ClaimsMappingPolicyDraft.From(_editing.Definition.Value)
      : new();
    _draft.DisplayName = Context.Original.Policy.DisplayName;
    _tenantName = tenants.CurrentTenant.HasValue ? tenants.CurrentTenant.Value.Name : "No tenant selected";
    tenants.TenantChanged += OnTenantChanged;
  }

  private void DraftChanged()
  {
    _review = null;
    _preview = null;
    _definitionChanged = false;
    _confirmed = false;
    _error = null;
  }

  private void Review()
  {
    if (IsDisabled)
      return;
    var request = new UpdateClaimsMappingPolicyRequest(_draft.DisplayName, _editing.CanEditDefinition ? _draft.ToRequest() : null);
    var patch = request.CreatePatch(Context.Original.Policy);
    if (patch.IsFailure)
    {
      _error = new(GraphOperationErrorType.InvalidInput, patch.Error);
      return;
    }
    _error = null;
    _confirmed = false;
    _review = patch.Value.HasValue ? request : null;
    _definitionChanged = patch.Value.HasValue && patch.Value.Value.Definition is not null;
    if (_definitionChanged && Assignments is null)
    {
      _error = new(GraphOperationErrorType.GraphFailure, "Assignments could not be retrieved. Reload before changing this policy definition.");
      _review = null;
    }
    _preview = _editing.CanEditDefinition
      ? ClaimsMappingPolicyDefinitionSerializer.Serialize(_draft.ToRequest().ValidateAndNormalize().Value, indented: true) : null;
  }

  private async Task SaveAsync()
  {
    if (IsDisabled || _review is null || (NeedsConfirmation && !_confirmed))
      return;
    _isSaving = true;
    _error = null;
    using var cancellation = new CancellationTokenSource();
    _save = cancellation;
    try
    {
      var result = await service.UpdateAsync(Context, _review with
      {
        ConfirmedAssignmentIds = NeedsConfirmation && _confirmed ? AssignmentIds : null
      }, cancellation.Token);
      if (_disposed || _invalidated)
        return;
      if (result.IsFailure)
      {
        _error = result.Error;
        _review = null;
        _confirmed = false;
        _requiresReload = result.Error.Type is not GraphOperationErrorType.InvalidInput;
        _invalidated = result.Error.Type is GraphOperationErrorType.TenantChanged or GraphOperationErrorType.NotFound;
        if (result.Error.Type == GraphOperationErrorType.NotFound)
          await OnMissing.InvokeAsync(result.Error);
        return;
      }
      _invalidated = true;
      await OnSaved.InvokeAsync(result.Value);
    }
    catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
    {
      // Tenant selection or navigation superseded this draft.
    }
    finally
    {
      _save = null;
      _isSaving = false;
    }
  }

  private void OnTenantChanged()
  {
    _invalidated = true;
    _save?.Cancel();
    _draft = new();
    _review = null;
    _preview = null;
    _error = new(GraphOperationErrorType.TenantChanged, "The tenant changed. Reopen the policy in the selected tenant.");
    _ = InvokeAsync(StateHasChanged);
  }

  public void Dispose()
  {
    _disposed = true;
    tenants.TenantChanged -= OnTenantChanged;
    _save?.Cancel();
  }
}
