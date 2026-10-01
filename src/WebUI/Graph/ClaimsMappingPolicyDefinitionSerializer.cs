using System.Text.Json;
using System.Text.Json.Serialization;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public static class ClaimsMappingPolicyDefinitionSerializer
{
  public static string Serialize(CreateClaimsMappingPolicyRequest validatedRequest, bool indented = false) =>
    JsonSerializer.Serialize(new DefinitionEnvelope(new PolicyDefinition(
      validatedRequest.IncludeBasicClaimSet ? "true" : "false",
      validatedRequest.ClaimsSchema.Select(claim => new ClaimDefinition(
        claim.Source, claim.Id, claim.JwtClaimType, claim.SamlClaimType, claim.Value)).ToArray())),
      new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, WriteIndented = indented });

  private sealed record DefinitionEnvelope(PolicyDefinition ClaimsMappingPolicy);

  // Microsoft Entra expects a string Boolean in the inner document, not a JSON Boolean.
  private sealed record PolicyDefinition(string IncludeBasicClaimSet, IReadOnlyList<ClaimDefinition> ClaimsSchema)
  {
    public int Version => 1;
  }

  private sealed record ClaimDefinition(string? Source, [property: JsonPropertyName("ID")] string? Id,
    string? JwtClaimType, string? SamlClaimType, string? Value);
}
