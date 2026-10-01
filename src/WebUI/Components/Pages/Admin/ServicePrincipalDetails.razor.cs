using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.AspNetCore.Components;
using DetailsModel = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph.ServicePrincipalDetails;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class ServicePrincipalDetails(ICurrentTenantContext tenants, IServicePrincipalService service,
  IApplicationService relatedService, NavigationManager navigation) : IDisposable
{
  [Parameter] public string ObjectId { get; set; } = "";
  private TenantViewState<Inspection> _state = default!;
  private GraphOperationError? DisplayError => _state.Error ??
    (_state.Value is { Related.IsFailure: true } inspection ? inspection.Related.Error : null);

  protected override void OnInitialized() => _state = new(tenants, OnTenantChanged);
  protected override Task OnParametersSetAsync() => LoadAsync();

  private Task LoadAsync() => _state.LoadAsync(async token =>
  {
    if (!Guid.TryParse(ObjectId, out var objectId) || objectId == Guid.Empty)
      return new GraphOperationError(GraphOperationErrorType.InvalidInput, "Service Principal Object ID must be a non-empty GUID.");
    var result = await service.GetByObjectIdAsync(objectId, token);
    if (result.IsFailure)
      return result.Error;
    token.ThrowIfCancellationRequested();
    // The relationship uses Client ID, never either directory Object ID.
    var related = result.Value.ServicePrincipal.AppId is { } appId
      ? await relatedService.FindByAppIdAsync(appId, token)
      : Result.Success<Maybe<Graph.ApplicationDetails>, GraphOperationError>(Maybe<Graph.ApplicationDetails>.None);
    return new Inspection(result.Value, related);
  });

  private void OpenRelated()
  {
    if (_state.Value is { Related.IsSuccess: true } inspection && inspection.Related.Value.HasValue)
      navigation.NavigateTo($"/applications/{inspection.Related.Value.Value.Application.ObjectId:D}");
  }

  private void OnTenantChanged() => _ = InvokeAsync(() =>
  {
    navigation.NavigateTo("/service-principals");
    StateHasChanged();
  });

  public void Dispose() => _state.Dispose();

  private sealed record Inspection(DetailsModel Details,
    Result<Maybe<Graph.ApplicationDetails>, GraphOperationError> Related);
}
