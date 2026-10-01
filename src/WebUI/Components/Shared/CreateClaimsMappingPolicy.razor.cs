using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class CreateClaimsMappingPolicy(ICurrentTenantContext tenants, IClaimsMappingPolicyService service) : IDisposable
{
  [Parameter, EditorRequired] public ClaimsMappingPolicyCreationContext Context { get; set; } = default!;
  [Parameter] public EventCallback<ClaimsMappingPolicyListItem> OnCreated { get; set; }
  [Parameter] public EventCallback OnCancel { get; set; }

  private string _displayName = "";
  private bool _includeBasicClaimSet = true;
  private readonly List<ClaimDraft> _claims = [];
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

  private void AddClaim()
  {
    if (IsDisabled || _claims.Count >= CreateClaimsMappingPolicyRequest.MaximumClaims)
      return;
    _claims.Add(new());
    DraftChanged();
  }

  private void RemoveClaim(ClaimDraft claim)
  {
    if (IsDisabled)
      return;
    _claims.Remove(claim);
    DraftChanged();
  }

  private void ModeChanged(ClaimDraft claim)
  {
    claim.Value = "";
    claim.Id = "";
    claim.Source = "user";
    DraftChanged();
  }

  private void SourceChanged(ClaimDraft claim)
  {
    claim.Id = "";
    DraftChanged();
  }

  private static Task<IEnumerable<string>> SearchPropertiesAsync(string source, string text, CancellationToken token)
  {
    token.ThrowIfCancellationRequested();
    return Task.FromResult(ClaimsMappingClaimSources.PropertiesFor(source)
      .Where(property => string.IsNullOrEmpty(text) || property.Contains(text, StringComparison.OrdinalIgnoreCase)));
  }

  private void Review()
  {
    if (IsDisabled)
      return;
    var request = new CreateClaimsMappingPolicyRequest(_displayName, _includeBasicClaimSet,
      _claims.Select(claim => new CreateClaimsSchemaEntry(claim.Mode, claim.Source, claim.Id,
        claim.JwtClaimType, claim.SamlClaimType, claim.Value)).ToArray()).ValidateAndNormalize();
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
    _claims.Clear();
    _displayName = "";
    _error = new(GraphOperationErrorType.TenantChanged, "The tenant changed. Cancel and start a new policy draft.");
    _ = InvokeAsync(StateHasChanged);
  }

  public void Dispose()
  {
    _disposed = true;
    tenants.TenantChanged -= OnTenantChanged;
    _creation?.Cancel();
  }

  private sealed class ClaimDraft
  {
    public ClaimValueMode Mode { get; set; }
    public string Source { get; set; } = "user";
    public string Id { get; set; } = "";
    public string JwtClaimType { get; set; } = "";
    public string SamlClaimType { get; set; } = "";
    public string Value { get; set; } = "";
  }
}
