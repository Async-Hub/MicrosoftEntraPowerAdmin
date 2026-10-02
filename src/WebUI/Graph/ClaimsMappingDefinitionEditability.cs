using System.Text.Json;
using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public enum ClaimsMappingDefinitionEditability
{
  FullySupported,
  UnsupportedDefinition,
  MalformedDefinition
}

public sealed record ClaimsMappingDefinitionEditing(
  ClaimsMappingDefinitionEditability Classification, string? Explanation,
  Maybe<CreateClaimsMappingPolicyRequest> Definition)
{
  public bool CanEditDefinition => Classification == ClaimsMappingDefinitionEditability.FullySupported;

  public static ClaimsMappingDefinitionEditing Inspect(ClaimsMappingPolicyListItem policy)
  {
    if (policy.Definitions.Count != 1)
      return Unsupported("The editor requires exactly one supported definition string.");
    var definition = policy.Definitions[0];
    try
    {
      using var original = JsonDocument.Parse(definition.Raw);
      var root = original.RootElement;
      if (!HasOnlyProperties(root, "ClaimsMappingPolicy") ||
          !root.TryGetProperty("ClaimsMappingPolicy", out var inner) ||
          !HasOnlyProperties(inner, "Version", "IncludeBasicClaimSet", "ClaimsSchema"))
        return Unsupported("This policy contains definition features (such as transformations or unknown properties) that the current editor cannot preserve.");
      if (definition.Document is not { Version: 1, IncludeBasicClaimSet: { } include, ClaimsSchema: { } schema } ||
          !inner.TryGetProperty("ClaimsSchema", out var claims) || claims.ValueKind != JsonValueKind.Array ||
          claims.EnumerateArray().Any(claim => !HasOnlyProperties(claim, "Source", "ID", "Value", "JwtClaimType", "SamlClaimType")))
        return Unsupported("The version or claims structure is not supported by the current editor.");

      var request = new CreateClaimsMappingPolicyRequest(policy.DisplayName, include,
        schema.Select(claim => new CreateClaimsSchemaEntry(claim.Value is null ? ClaimValueMode.DirectoryAttribute : ClaimValueMode.ConstantValue,
          claim.Source, claim.Id, claim.JwtClaimType, claim.SamlClaimType, claim.Value)).ToArray()).ValidateAndNormalize();
      if (request.IsFailure)
        return Unsupported("The existing claims cannot be represented using the editor's validation rules.");
      using var serialized = JsonDocument.Parse(ClaimsMappingPolicyDefinitionSerializer.Serialize(request.Value));
      // Check the entire original tree, including value types and optional properties. The tolerant
      // read parser alone cannot tell us whether serialization would discard configuration.
      if (!JsonElement.DeepEquals(root, serialized.RootElement))
        return Unsupported("The current editor cannot round-trip every existing definition value without changing it.");
      return new(ClaimsMappingDefinitionEditability.FullySupported, null, request.Value);
    }
    catch (JsonException)
    {
      return new(ClaimsMappingDefinitionEditability.MalformedDefinition,
        "MEPA could not safely parse this policy definition. The policy name can still be changed, but definition editing is disabled.",
        Maybe<CreateClaimsMappingPolicyRequest>.None);
    }
  }

  private static bool HasOnlyProperties(JsonElement element, params string[] allowed)
  {
    if (element.ValueKind != JsonValueKind.Object)
      return false;
    var seen = new HashSet<string>(StringComparer.Ordinal);
    return element.EnumerateObject().All(property => allowed.Contains(property.Name, StringComparer.Ordinal) && seen.Add(property.Name));
  }

  private static ClaimsMappingDefinitionEditing Unsupported(string reason) => new(
    ClaimsMappingDefinitionEditability.UnsupportedDefinition,
    reason + " You can rename the policy, but definition editing is disabled to prevent losing existing configuration.",
    Maybe<CreateClaimsMappingPolicyRequest>.None);
}
