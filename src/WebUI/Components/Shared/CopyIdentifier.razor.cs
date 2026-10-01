using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using MudBlazor;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class CopyIdentifier(IJSRuntime javascript, ISnackbar snackbar)
{
  [Parameter] public Guid? Value { get; set; }
  [Parameter, EditorRequired] public string Label { get; set; } = "Identifier";

  private async Task CopyAsync()
  {
    if (Value is not { } value)
      return;
    try
    {
      await javascript.InvokeVoidAsync("navigator.clipboard.writeText", value.ToString("D"));
      snackbar.Add($"{Label} copied.", Severity.Success);
    }
    catch (JSDisconnectedException)
    {
      // The circuit has closed; there is no remaining UI to notify.
    }
    catch (JSException)
    {
      snackbar.Add("Clipboard access is unavailable. Select and copy the identifier manually.", Severity.Warning);
    }
  }
}
