using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.AspNetCore.Components;
using DetailsModel = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph.ApplicationDetails;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class ApplicationDetails(ICurrentTenantContext tenants, IApplicationService service,
  IServicePrincipalService relatedService, NavigationManager navigation) : IDisposable
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
      return new GraphOperationError(GraphOperationErrorType.InvalidInput, "Application Object ID must be a non-empty GUID.");
    var result = await service.GetByObjectIdAsync(objectId, token);
    if (result.IsFailure)
      return result.Error;
    token.ThrowIfCancellationRequested();
    // The relationship uses Client ID, never either directory Object ID.
    var related = result.Value.Application.AppId is { } appId
      ? await relatedService.FindByAppIdAsync(appId, token)
      : Result.Success<Maybe<Graph.ServicePrincipalDetails>, GraphOperationError>(Maybe<Graph.ServicePrincipalDetails>.None);
    return new Inspection(result.Value, related);
  });

  private void OpenRelated()
  {
    if (_state.Value is { Related.IsSuccess: true } inspection && inspection.Related.Value.HasValue)
      navigation.NavigateTo($"/service-principals/{inspection.Related.Value.Value.ServicePrincipal.ObjectId:D}");
  }

  private void OnTenantChanged() => _ = InvokeAsync(() =>
  {
    navigation.NavigateTo("/applications");
    StateHasChanged();
  });

  public void Dispose() => _state.Dispose();

  private sealed record Inspection(DetailsModel Details,
    Result<Maybe<Graph.ServicePrincipalDetails>, GraphOperationError> Related);
}
