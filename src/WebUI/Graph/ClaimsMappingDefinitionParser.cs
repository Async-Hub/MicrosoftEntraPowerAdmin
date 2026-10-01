using System.Text.Json;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public static class ClaimsMappingDefinitionParser
{
  public static ClaimsMappingDefinition Parse(string raw)
  {
    if (string.IsNullOrWhiteSpace(raw))
      return new(raw, raw, null, ["The definition is empty."]);

    try
    {
      using var json = JsonDocument.Parse(raw);
      // Format the original JSON tree, never the understood subset of the document.
      var formatted = JsonSerializer.Serialize(json.RootElement, new JsonSerializerOptions { WriteIndented = true });
      if (json.RootElement.ValueKind != JsonValueKind.Object ||
          !json.RootElement.TryGetProperty("ClaimsMappingPolicy", out var policy) || policy.ValueKind != JsonValueKind.Object)
        return new(raw, formatted, null, ["This JSON structure is not understood. Inspect the raw definition."]);

      var warnings = new List<string>();
      int? version = null;
      if (policy.TryGetProperty("Version", out var versionJson) && versionJson.ValueKind != JsonValueKind.Null)
      {
        if (versionJson.ValueKind == JsonValueKind.Number && versionJson.TryGetInt32(out var number))
          version = number;
        else warnings.Add("Version is not an integer.");
      }
      bool? includeBasicClaimSet = null;
      if (policy.TryGetProperty("IncludeBasicClaimSet", out var include) && include.ValueKind != JsonValueKind.Null)
      {
        if (include.ValueKind is JsonValueKind.True or JsonValueKind.False)
          includeBasicClaimSet = include.GetBoolean();
        else if (include.ValueKind == JsonValueKind.String && bool.TryParse(include.GetString(), out var boolean))
          includeBasicClaimSet = boolean;
        else warnings.Add("IncludeBasicClaimSet is not a Boolean or Boolean string.");
      }

      var schema = Entries(policy, "ClaimsSchema", warnings, entry => new ClaimsSchemaEntry(
        Text(entry, "Source", warnings), Text(entry, "ID", warnings), Text(entry, "JwtClaimType", warnings),
        Text(entry, "SamlClaimType", warnings), Text(entry, "Value", warnings), Text(entry, "TransformationId", warnings)));
      // Graph examples contain both singular and plural spellings.
      var transformationName = policy.TryGetProperty("ClaimsTransformation", out _) ? "ClaimsTransformation" : "ClaimsTransformations";
      var transformations = Entries(policy, transformationName, warnings, entry => new ClaimsTransformationEntry(
        Text(entry, "ID", warnings), Text(entry, "TransformationMethod", warnings),
        Entries(entry, "InputClaims", warnings, claim => ReadClaim(claim, warnings)),
        Entries(entry, "InputParameters", warnings, parameter => new TransformationParameter(
          Text(parameter, "ID", warnings), Text(parameter, "DataType", warnings), Text(parameter, "Value", warnings))),
        Entries(entry, "OutputClaims", warnings, claim => ReadClaim(claim, warnings))));

      return new(raw, formatted, new(version, includeBasicClaimSet, schema, transformations),
        Array.AsReadOnly(warnings.Distinct().ToArray()));
    }
    catch (JsonException)
    {
      return new(raw, raw, null, ["The definition contains malformed JSON or exceeds the supported JSON depth. Its original text is shown below."]);
    }
  }

  private static TransformationClaim ReadClaim(JsonElement claim, List<string> warnings) => new(
    Text(claim, "ClaimTypeReferenceId", warnings), Text(claim, "TransformationClaimType", warnings));

  private static IReadOnlyList<T>? Entries<T>(JsonElement parent, string name, List<string> warnings, Func<JsonElement, T> read)
  {
    if (!parent.TryGetProperty(name, out var array) || array.ValueKind == JsonValueKind.Null)
      return null;
    if (array.ValueKind != JsonValueKind.Array)
    {
      warnings.Add($"{name} is not an array. Inspect the raw definition.");
      return null;
    }
    var entries = new List<T>();
    foreach (var element in array.EnumerateArray())
    {
      if (element.ValueKind == JsonValueKind.Object)
        entries.Add(read(element));
      else warnings.Add($"Some {name} entries are not objects and cannot be displayed in the structured view.");
    }
    return entries.AsReadOnly();
  }

  private static string? Text(JsonElement parent, string name, List<string> warnings)
  {
    if (!parent.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
      return null;
    if (value.ValueKind == JsonValueKind.String)
      return value.GetString();
    if (value.ValueKind is JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False)
      return value.GetRawText();
    warnings.Add($"{name} has an unsupported value. Inspect the raw definition.");
    return null;
  }
}
