namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

// Mutable UI state is owned by one editor, never by a cached Graph policy.
public sealed class ClaimsMappingPolicyDraft
{
  public string DisplayName { get; set; } = "";
  public bool IncludeBasicClaimSet { get; set; } = true;
  public List<ClaimDraft> Claims { get; } = [];

  public CreateClaimsMappingPolicyRequest ToRequest() => new(DisplayName, IncludeBasicClaimSet,
    Claims.Select(claim => new CreateClaimsSchemaEntry(claim.Mode, claim.Source, claim.Id,
      claim.JwtClaimType, claim.SamlClaimType, claim.Value)).ToArray());

  public static ClaimsMappingPolicyDraft From(CreateClaimsMappingPolicyRequest request)
  {
    var draft = new ClaimsMappingPolicyDraft { DisplayName = request.DisplayName, IncludeBasicClaimSet = request.IncludeBasicClaimSet };
    draft.Claims.AddRange(request.ClaimsSchema.Select(claim => new ClaimDraft
    {
      Mode = claim.Mode,
      Source = claim.Source ?? "user",
      Id = claim.Id ?? "",
      JwtClaimType = claim.JwtClaimType ?? "",
      SamlClaimType = claim.SamlClaimType ?? "",
      Value = claim.Value ?? ""
    }));
    return draft;
  }

  public sealed class ClaimDraft
  {
    public ClaimValueMode Mode { get; set; }
    public string Source { get; set; } = "user";
    public string Id { get; set; } = "";
    public string JwtClaimType { get; set; } = "";
    public string SamlClaimType { get; set; } = "";
    public string Value { get; set; } = "";
  }
}
