using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class ClaimsMappingPolicyDefinitionEditor
{
  [Parameter, EditorRequired] public ClaimsMappingPolicyDraft Draft { get; set; } = default!;
  [Parameter] public bool IsDisabled { get; set; }
  [Parameter] public EventCallback OnChanged { get; set; }

  private Task DraftChanged() => OnChanged.InvokeAsync();

  private Task AddClaim()
  {
    if (IsDisabled || Draft.Claims.Count >= CreateClaimsMappingPolicyRequest.MaximumClaims)
      return Task.CompletedTask;
    Draft.Claims.Add(new());
    return DraftChanged();
  }

  private Task RemoveClaim(ClaimsMappingPolicyDraft.ClaimDraft claim)
  {
    if (IsDisabled)
      return Task.CompletedTask;
    Draft.Claims.Remove(claim);
    return DraftChanged();
  }

  private Task ModeChanged(ClaimsMappingPolicyDraft.ClaimDraft claim)
  {
    claim.Value = "";
    claim.Id = "";
    claim.Source = "user";
    return DraftChanged();
  }

  private Task SourceChanged(ClaimsMappingPolicyDraft.ClaimDraft claim)
  {
    claim.Id = "";
    return DraftChanged();
  }

  private static Task<IEnumerable<string>> SearchPropertiesAsync(string source, string text, CancellationToken token)
  {
    token.ThrowIfCancellationRequested();
    return Task.FromResult(ClaimsMappingClaimSources.PropertiesFor(source)
      .Where(property => string.IsNullOrEmpty(text) || property.Contains(text, StringComparison.OrdinalIgnoreCase)));
  }
}
