namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed record ClaimsMappingPolicyDeletionContext(
  Guid TenantId, Guid SelectionId, ClaimsMappingPolicyDetails Original, string? CurrentDisplayName)
{
  public string ConfirmationPhrase => PhraseFor(CurrentDisplayName);
  public bool HasAssignments => Original.Assignments.IsSuccess && AssignmentCount > 0;
  public int? AssignmentCount => Original.Assignments.IsSuccess
    ? Original.Assignments.Value.ServicePrincipals.Count + Original.Assignments.Value.Applications.Count
      + Original.Assignments.Value.OtherObjects.Count : null;
  public bool CanConfirm(string? entered) => AssignmentCount == 0 && Matches(CurrentDisplayName, entered);

  public static string PhraseFor(string? displayName) => string.IsNullOrWhiteSpace(displayName) || displayName.Any(char.IsControl)
    ? "DELETE" : displayName.Trim();

  public static bool Matches(string? displayName, string? entered) =>
    string.Equals(PhraseFor(displayName), entered?.Trim(), StringComparison.Ordinal);
}
