using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class ClaimsMappingPolicyTable
{
  [Parameter, EditorRequired] public IReadOnlyList<ClaimsMappingPolicyListItem> Policies { get; set; } = [];
  [Parameter] public EventCallback<Guid> OnOpen { get; set; }
  [Parameter] public string EmptyMessage { get; set; } = "No Claims Mapping Policies in the current tenant.";
}
