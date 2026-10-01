using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed record ClaimsMappingPolicyListItem(
  Guid ObjectId, string DisplayName, bool? IsOrganizationDefault,
  IReadOnlyList<ClaimsMappingDefinition> Definitions)
{
  public bool HasDefinition => Definitions.Any(definition => !string.IsNullOrWhiteSpace(definition.Raw));
  public int? ClaimsSchemaCount => CountEntries(document => document.ClaimsSchema?.Count);
  public int? ClaimsTransformationCount => CountEntries(document => document.ClaimsTransformation?.Count);

  private int? CountEntries(Func<ClaimsMappingDocument, int?> count)
  {
    if (Definitions.Count == 0)
      return null;
    var total = 0;
    foreach (var definition in Definitions)
    {
      if (definition.Document is not { } document || count(document) is not { } entries)
        return null;
      total += entries;
    }
    return total;
  }
}

public sealed record ClaimsMappingPolicyDetails(ClaimsMappingPolicyListItem Policy,
  Result<ClaimsMappingAssignments, GraphOperationError> Assignments);

public sealed record ClaimsMappingAssignments(
  IReadOnlyList<ServicePrincipalListItem> ServicePrincipals,
  IReadOnlyList<ApplicationListItem> Applications,
  IReadOnlyList<PolicyDirectoryObject> OtherObjects);

public sealed record PolicyDirectoryObject(Guid ObjectId, string? Type);

public sealed record ClaimsMappingDefinition(string Raw, string Formatted,
  ClaimsMappingDocument? Document, IReadOnlyList<string> Warnings);

public sealed record ClaimsMappingDocument(int? Version, bool? IncludeBasicClaimSet,
  IReadOnlyList<ClaimsSchemaEntry>? ClaimsSchema,
  IReadOnlyList<ClaimsTransformationEntry>? ClaimsTransformation);

public sealed record ClaimsSchemaEntry(string? Source, string? Id, string? JwtClaimType,
  string? SamlClaimType, string? Value, string? TransformationId);

public sealed record ClaimsTransformationEntry(string? Id, string? TransformationMethod,
  IReadOnlyList<TransformationClaim>? InputClaims,
  IReadOnlyList<TransformationParameter>? InputParameters,
  IReadOnlyList<TransformationClaim>? OutputClaims);

public sealed record TransformationClaim(string? ClaimTypeReferenceId, string? TransformationClaimType);
public sealed record TransformationParameter(string? Id, string? DataType, string? Value);
