using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;
using MudBlazor;
using DetailsModel = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph.ClaimsMappingPolicyDetails;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class ClaimsMappingPolicyDetails(ICurrentTenantContext tenants, IClaimsMappingPolicyService service,
  NavigationManager navigation, ISnackbar snackbar) : IDisposable
{
  [Parameter] public string ObjectId { get; set; } = "";
  private TenantViewState<DetailsModel> _state = default!;
  private bool _tenantChanged;
  private ClaimsMappingPolicyAssignmentContext? _confirmation;
  private ClaimsMappingPolicyListItem? _confirmationPolicy;
  private ServicePrincipalListItem? _unassignPrincipal;
  private IReadOnlyList<Guid> _assignedPrincipalIds = [];
  private GraphOperationError? _operationError;
  private GraphOperationError? DisplayError => _state.Error ??
    (_state.Value is { Assignments.IsFailure: true } details ? details.Assignments.Error : null);
  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);
  protected override Task OnParametersSetAsync()
  {
    Cancel();
    _operationError = null;
    return LoadAsync();
  }

  private Task LoadAsync() => _tenantChanged ? Task.CompletedTask : _state.LoadAsync(token =>
    Guid.TryParse(ObjectId, out var objectId) && objectId != Guid.Empty
      ? service.GetByObjectIdAsync(objectId, token)
      : Task.FromResult(CSharpFunctionalExtensions.Result.Failure<DetailsModel, GraphOperationError>(
        new(GraphOperationErrorType.InvalidInput, "Policy Object ID must be a non-empty GUID."))));

  private void Assign() => BeginConfirmation(null);
  private void Unassign(ServicePrincipalListItem principal) => BeginConfirmation(principal);

  private void BeginConfirmation(ServicePrincipalListItem? principal)
  {
    if (_tenantChanged || _state.IsLoading || _confirmation is not null || _state.Value is not { Assignments.IsSuccess: true } details)
      return;
    if (principal is not null && !details.Assignments.Value.ServicePrincipals.Any(item => item.ObjectId == principal.ObjectId))
      return;
    var started = service.BeginAssignment();
    _operationError = started.IsFailure ? started.Error : null;
    if (started.IsFailure)
      return;
    _confirmation = started.Value;
    _confirmationPolicy = details.Policy;
    _unassignPrincipal = principal;
    _assignedPrincipalIds = details.Assignments.Value.ServicePrincipals.Select(item => item.ObjectId).ToArray();
  }

  private void Cancel()
  {
    _confirmation = null;
    _confirmationPolicy = null;
    _unassignPrincipal = null;
    _assignedPrincipalIds = [];
  }

  private async Task FinishedAsync(GraphOperationError? error)
  {
    var unassigned = _unassignPrincipal is not null;
    Cancel();
    if (_tenantChanged)
      return;
    _operationError = error;
    // Reloads policy-side appliesTo, preserving its direction and any independently reported read error.
    await LoadAsync();
    if (!_tenantChanged && error is null)
      snackbar.Add(unassigned ? "Policy unassigned." : "Policy assigned.", Severity.Success);
  }

  private void OpenServicePrincipal(Guid objectId)
  {
    if (_state.Value is { Assignments.IsSuccess: true } details && details.Assignments.Value.ServicePrincipals.Any(item => item.ObjectId == objectId))
      navigation.NavigateTo($"/service-principals/{objectId:D}");
  }

  private void OpenApplication(Guid objectId)
  {
    if (_state.Value is { Assignments.IsSuccess: true } details && details.Assignments.Value.Applications.Any(item => item.ObjectId == objectId))
      navigation.NavigateTo($"/applications/{objectId:D}");
  }

  private void OnTenantChanged()
  {
    _tenantChanged = true;
    Cancel();
    _operationError = null;
    _ = InvokeAsync(() =>
    {
      navigation.NavigateTo("/claims-mapping-policies");
      StateHasChanged();
    });
  }

  public void Dispose() => _state.Dispose();
}
