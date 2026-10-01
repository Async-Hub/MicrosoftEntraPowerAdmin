using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class ChangeClaimsMappingPolicyAssignment(ICurrentTenantContext tenants,
  IClaimsMappingPolicyService policies, IServicePrincipalService principals) : IDisposable
{
  [Parameter, EditorRequired] public ClaimsMappingPolicyAssignmentContext Context { get; set; } = default!;
  [Parameter] public ServicePrincipalListItem? Principal { get; set; }
  [Parameter] public ClaimsMappingPolicyListItem? Policy { get; set; }
  [Parameter] public bool IsUnassignment { get; set; }
  [Parameter] public IReadOnlyList<Guid> AssignedPrincipalIds { get; set; } = [];
  [Parameter] public EventCallback<GraphOperationError?> OnFinished { get; set; }
  [Parameter] public EventCallback OnCancel { get; set; }

  private readonly CancellationTokenSource _lifetime = new();
  private ServicePrincipalListItem? _principal;
  private ClaimsMappingPolicyListItem? _policy;
  private IReadOnlyList<ClaimsMappingPolicyListItem> _availablePolicies = [];
  private IReadOnlyList<ClaimsMappingPolicyListItem> _assignedPolicies = [];
  private DirectoryPage<ServicePrincipalListItem>? _principalPage;
  private string _search = "";
  private string _policyFilter = "";
  private string _tenantName = "";
  private GraphOperationError? _error;
  private bool _loading;
  private bool _mutating;
  private bool _invalidated;
  private bool _disposed;
  private bool IsBusy => _loading || _mutating || _invalidated;
  private bool CanConfirm => !IsBusy && _error is null && _principal is not null && _policy is not null;
  private string ActionLabel => IsUnassignment ? "Unassign" : "Assign";
  private static string PolicyLabel(ClaimsMappingPolicyListItem? policy) => policy is null
    ? "" : $"{policy.DisplayName} — Policy Object ID: {policy.ObjectId:D}";
  private IEnumerable<ClaimsMappingPolicyListItem> FilteredPolicies => _availablePolicies.Where(item =>
    item.DisplayName.Contains(_policyFilter, StringComparison.OrdinalIgnoreCase)
    || item.ObjectId.ToString().Contains(_policyFilter, StringComparison.OrdinalIgnoreCase));
  private IEnumerable<ServicePrincipalListItem> AvailablePrincipals => _principalPage?.Items.Where(item =>
    !AssignedPrincipalIds.Contains(item.ObjectId)) ?? [];

  protected override async Task OnInitializedAsync()
  {
    tenants.TenantChanged += OnTenantChanged;
    if (tenants.CurrentTenant.HasNoValue || tenants.CurrentTenant.Value.TenantId != Context.TenantId)
    {
      Invalidate();
      return;
    }
    _tenantName = tenants.CurrentTenant.Value.Name;
    _principal = Principal;
    _policy = Policy;
    if (IsUnassignment || Principal is null)
    {
      if (!IsUnassignment)
        await SearchAsync();
      return;
    }
    await LoadAsync(async token =>
    {
      if (!await LoadAssignedAsync(token))
        return;
      var result = await policies.ListAsync(token);
      if (result.IsFailure)
        _error = result.Error;
      else
        _availablePolicies = result.Value.Where(item => !_assignedPolicies.Any(assigned => assigned.ObjectId == item.ObjectId)).ToArray();
    });
  }

  private Task SearchAsync() => SearchPageAsync(null);
  private Task NextPageAsync() => SearchPageAsync(_principalPage?.NextPage);
  private Task SearchPageAsync(DirectoryContinuation? next) => LoadAsync(async token =>
  {
    _principal = null;
    _assignedPolicies = [];
    _principalPage = null;
    var result = await principals.SearchAsync(_search, next, token);
    if (result.IsFailure)
      _error = result.Error;
    else
      _principalPage = result.Value;
  });

  private Task SelectPrincipalAsync(ServicePrincipalListItem item) => LoadAsync(async token =>
  {
    if (!AvailablePrincipals.Any(candidate => candidate.ObjectId == item.ObjectId))
      return;
    _principal = item;
    if (await LoadAssignedAsync(token) && _policy is { } policy && _assignedPolicies.Any(assigned => assigned.ObjectId == policy.ObjectId))
      _error = new(GraphOperationErrorType.InvalidInput, "This policy is already assigned to the selected Enterprise Application. Choose another application.");
  });

  private async Task<bool> LoadAssignedAsync(CancellationToken token)
  {
    if (_principal is null)
      return false;
    var result = await policies.GetAssignedPoliciesAsync(Context, _principal.ObjectId, token);
    if (result.IsFailure)
    {
      _error = result.Error;
      return false;
    }
    _assignedPolicies = result.Value;
    return true;
  }

  private async Task LoadAsync(Func<CancellationToken, Task> load)
  {
    if (IsBusy)
      return;
    _loading = true;
    _error = null;
    try
    {
      await load(_lifetime.Token);
    }
    catch (OperationCanceledException) when (_lifetime.IsCancellationRequested)
    {
      // Tenant changes and navigation supersede the selection.
    }
    finally
    {
      _loading = false;
      if (_invalidated || _disposed)
        ClearSelections();
    }
  }

  private async Task ConfirmAsync()
  {
    if (!CanConfirm || _principal is not { } principal || _policy is not { } policy)
      return;
    _mutating = true;
    try
    {
      var result = IsUnassignment
        ? await policies.UnassignAsync(Context, principal.ObjectId, policy.ObjectId, _lifetime.Token)
        : await policies.AssignAsync(Context, principal.ObjectId, policy.ObjectId, _lifetime.Token);
      if (_invalidated || _disposed)
        return;
      // Close this confirmation after every attempt; the parent reloads authoritative Graph state.
      _invalidated = true;
      await OnFinished.InvokeAsync(result.IsFailure ? result.Error : null);
    }
    catch (OperationCanceledException) when (_lifetime.IsCancellationRequested)
    {
      // A submitted operation may finish in the original tenant; never display it in another tenant.
    }
    finally
    {
      _mutating = false;
    }
  }

  private void ClearSelections()
  {
    _principal = null;
    _policy = null;
    _availablePolicies = [];
    _assignedPolicies = [];
    _principalPage = null;
    _search = "";
    _policyFilter = "";
  }

  private void Invalidate()
  {
    _invalidated = true;
    _lifetime.Cancel();
    ClearSelections();
    _error = new(GraphOperationErrorType.TenantChanged, "The tenant changed. Cancel and start a new assignment confirmation.");
  }

  private void OnTenantChanged()
  {
    Invalidate();
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
