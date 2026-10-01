using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;
using MudBlazor;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class AssignedClaimsMappingPolicies(ICurrentTenantContext tenants, IClaimsMappingPolicyService service,
  NavigationManager navigation, ISnackbar snackbar) : IDisposable
{
  [Parameter, EditorRequired] public ServicePrincipalListItem ServicePrincipal { get; set; } = default!;
  private TenantViewState<IReadOnlyList<ClaimsMappingPolicyListItem>> _state = default!;
  private bool _tenantChanged;
  private ClaimsMappingPolicyAssignmentContext? _context;
  private ClaimsMappingPolicyAssignmentContext? _confirmation;
  private ClaimsMappingPolicyListItem? _unassignPolicy;
  private GraphOperationError? _operationError;
  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);
  protected override Task OnParametersSetAsync() => LoadAsync();
  private async Task LoadAsync()
  {
    if (_tenantChanged)
      return;
    var started = service.BeginAssignment();
    _context = started.IsSuccess ? started.Value : null;
    await _state.LoadAsync(token => started.IsSuccess
      ? service.GetAssignedPoliciesAsync(started.Value, ServicePrincipal.ObjectId, token)
      : Task.FromResult(CSharpFunctionalExtensions.Result.Failure<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>(started.Error)));
  }

  private void Assign()
  {
    if (_tenantChanged || _state.IsLoading || _state.Value is null || _confirmation is not null)
      return;
    _operationError = null;
    _unassignPolicy = null;
    _confirmation = _context;
  }

  private void Unassign(ClaimsMappingPolicyListItem policy)
  {
    if (_tenantChanged || _state.IsLoading || _confirmation is not null || _state.Value?.Any(item => item.ObjectId == policy.ObjectId) != true)
      return;
    _operationError = null;
    _unassignPolicy = policy;
    _confirmation = _context;
  }

  private void Cancel()
  {
    _confirmation = null;
    _unassignPolicy = null;
  }

  private async Task FinishedAsync(GraphOperationError? error)
  {
    var unassigned = _unassignPolicy is not null;
    Cancel();
    if (_tenantChanged)
      return;
    _operationError = error;
    await LoadAsync();
    if (!_tenantChanged && error is null)
      snackbar.Add(unassigned ? "Policy unassigned." : "Policy assigned.", Severity.Success);
  }

  private void Open(Guid objectId)
  {
    if (_state.Value?.Any(item => item.ObjectId == objectId) == true)
      navigation.NavigateTo($"/claims-mapping-policies/{objectId:D}");
  }

  private void OnTenantChanged()
  {
    _tenantChanged = true;
    Cancel();
    _context = null;
    _operationError = null;
    _ = InvokeAsync(StateHasChanged);
  }

  public void Dispose() => _state.Dispose();
}
