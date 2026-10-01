using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;
using DetailsModel = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph.ClaimsMappingPolicyDetails;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class ClaimsMappingPolicyDetails(ICurrentTenantContext tenants, IClaimsMappingPolicyService service,
  NavigationManager navigation) : IDisposable
{
  [Parameter] public string ObjectId { get; set; } = "";
  private TenantViewState<DetailsModel> _state = default!;
  private bool _tenantChanged;
  private GraphOperationError? DisplayError => _state.Error ??
    (_state.Value is { Assignments.IsFailure: true } details ? details.Assignments.Error : null);
  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);
  protected override Task OnParametersSetAsync() => LoadAsync();

  private Task LoadAsync() => _tenantChanged ? Task.CompletedTask : _state.LoadAsync(token =>
    Guid.TryParse(ObjectId, out var objectId) && objectId != Guid.Empty
      ? service.GetByObjectIdAsync(objectId, token)
      : Task.FromResult(CSharpFunctionalExtensions.Result.Failure<DetailsModel, GraphOperationError>(
        new(GraphOperationErrorType.InvalidInput, "Policy Object ID must be a non-empty GUID."))));

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
    _ = InvokeAsync(() =>
    {
      navigation.NavigateTo("/claims-mapping-policies");
      StateHasChanged();
    });
  }

  public void Dispose() => _state.Dispose();
}
