using CSharpFunctionalExtensions;
using Microsoft.Graph.Models;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed record ClaimsMappingPolicyEditContext(Guid TenantId, Guid SelectionId, ClaimsMappingPolicyDetails Original);

public sealed record UpdateClaimsMappingPolicyRequest(string DisplayName,
  CreateClaimsMappingPolicyRequest? Definition = null, IReadOnlyList<Guid>? ConfirmedAssignmentIds = null)
{
  public Result<Maybe<ClaimsMappingPolicy>> CreatePatch(ClaimsMappingPolicyListItem original)
  {
    if (string.IsNullOrWhiteSpace(DisplayName))
      return Result.Failure<Maybe<ClaimsMappingPolicy>>("Display Name is required.");
    var name = DisplayName.Trim();
    var patch = new ClaimsMappingPolicy();
    var nameChanged = !string.Equals(DisplayName, original.DisplayName, StringComparison.Ordinal)
      && !string.Equals(name, original.DisplayName, StringComparison.Ordinal);
    if (nameChanged)
      patch.DisplayName = name;
    if (Definition is { } definition)
    {
      var editing = ClaimsMappingDefinitionEditing.Inspect(original);
      if (!editing.CanEditDefinition)
        return Result.Failure<Maybe<ClaimsMappingPolicy>>(editing.Explanation ?? "Definition editing is disabled.");
      var validated = (definition with { DisplayName = name }).ValidateAndNormalize();
      if (validated.IsFailure)
        return Result.Failure<Maybe<ClaimsMappingPolicy>>(validated.Error);
      var serialized = ClaimsMappingPolicyDefinitionSerializer.Serialize(validated.Value);
      if (serialized != ClaimsMappingPolicyDefinitionSerializer.Serialize(editing.Definition.Value))
        patch.Definition = [serialized];
    }
    return nameChanged || patch.Definition is not null ? Maybe<ClaimsMappingPolicy>.From(patch) : Maybe<ClaimsMappingPolicy>.None;
  }
}
