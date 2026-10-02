using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin;

public partial class ApplicationDetailValue
{
  [Parameter] public string? Value { get; set; }
  [Parameter] public bool? BooleanValue { get; set; }
  [Parameter] public IReadOnlyList<string>? Values { get; set; }
  [Parameter] public bool IsTechnical { get; set; }

  private string? ScalarValue => BooleanValue switch
  {
    true => "Yes",
    false => "No",
    null => Value
  };

  private string? ValueClass(string? value) => string.IsNullOrWhiteSpace(value)
    ? "application-detail-empty"
    : IsTechnical ? "directory-id" : null;

  private static string DisplayValue(string? value) =>
    string.IsNullOrWhiteSpace(value) ? "Not configured" : value;
}
