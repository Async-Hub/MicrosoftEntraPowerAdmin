using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using Microsoft.AspNetCore.Components;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class ClaimsMappingDefinitionView
{
  [Parameter, EditorRequired] public ClaimsMappingDefinition Definition { get; set; } = default!;
  private bool HasSource => Any(entry => entry.Source);
  private bool HasId => Any(entry => entry.Id);
  private bool HasJwtClaimType => Any(entry => entry.JwtClaimType);
  private bool HasSamlClaimType => Any(entry => entry.SamlClaimType);
  private bool HasValue => Any(entry => entry.Value);
  private bool HasTransformationId => Any(entry => entry.TransformationId);
  private bool HasSchemaFields => HasSource || HasId || HasJwtClaimType || HasSamlClaimType || HasValue || HasTransformationId;
  private bool Any(Func<ClaimsSchemaEntry, string?> field) =>
    Definition.Document?.ClaimsSchema?.Any(entry => field(entry) is not null) == true;

  private static IReadOnlyList<string> ClaimValues(IReadOnlyList<TransformationClaim>? claims) =>
    claims?.Select(claim => $"Claim Type Reference ID: {claim.ClaimTypeReferenceId ?? "Not available"}; Transformation Claim Type: {claim.TransformationClaimType ?? "Not available"}").ToArray() ?? [];
  private static IReadOnlyList<string> ParameterValues(IReadOnlyList<TransformationParameter>? parameters) =>
    parameters?.Select(parameter => $"ID: {parameter.Id ?? "Not available"}; Data Type: {parameter.DataType ?? "Not available"}; Value: {parameter.Value ?? "Not available"}").ToArray() ?? [];
}
