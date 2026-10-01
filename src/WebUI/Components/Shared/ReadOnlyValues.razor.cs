using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class ReadOnlyValues
{
  [Parameter, EditorRequired] public string Label { get; set; } = "";
  [Parameter, EditorRequired] public IReadOnlyList<string> Values { get; set; } = [];
}
