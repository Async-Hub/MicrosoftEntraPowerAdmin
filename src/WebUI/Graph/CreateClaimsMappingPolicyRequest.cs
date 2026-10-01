using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

// Selection identity also invalidates drafts after switching away and back to the same tenant.
public sealed record ClaimsMappingPolicyCreationContext(Guid TenantId, Guid SelectionId);

public enum ClaimValueMode
{
  DirectoryAttribute,
  ConstantValue
}

public sealed record CreateClaimsSchemaEntry(ClaimValueMode Mode, string? Source, string? Id,
  string? JwtClaimType, string? SamlClaimType, string? Value);

public sealed record CreateClaimsMappingPolicyRequest(string DisplayName, bool IncludeBasicClaimSet,
  IReadOnlyList<CreateClaimsSchemaEntry> ClaimsSchema)
{
  public const int MaximumClaims = 50;

  public Result<CreateClaimsMappingPolicyRequest> ValidateAndNormalize()
  {
    if (string.IsNullOrWhiteSpace(DisplayName))
      return Result.Failure<CreateClaimsMappingPolicyRequest>("Display Name is required.");
    if (ClaimsSchema.Count > MaximumClaims)
      return Result.Failure<CreateClaimsMappingPolicyRequest>("A policy can contain at most 50 claims.");

    var claims = new List<CreateClaimsSchemaEntry>();
    for (var index = 0; index < ClaimsSchema.Count; index++)
    {
      var claim = ClaimsSchema[index];
      var jwt = Optional(claim.JwtClaimType);
      var saml = Optional(claim.SamlClaimType);
      var prefix = $"Claim {index + 1}: ";
      if (jwt is null && saml is null)
        return Result.Failure<CreateClaimsMappingPolicyRequest>(prefix + "enter a JWT or SAML claim type.");
      switch (claim.Mode)
      {
        case ClaimValueMode.DirectoryAttribute:
          var source = Optional(claim.Source);
          var id = Optional(claim.Id);
          if (!ClaimsMappingClaimSources.IsSupported(source))
            return Result.Failure<CreateClaimsMappingPolicyRequest>(prefix + "select a supported directory source.");
          if (id is null)
            return Result.Failure<CreateClaimsMappingPolicyRequest>(prefix + "attribute ID is required.");
          claims.Add(claim with { Source = source, Id = id, JwtClaimType = jwt, SamlClaimType = saml, Value = null });
          break;
        case ClaimValueMode.ConstantValue:
          if (string.IsNullOrWhiteSpace(claim.Value))
            return Result.Failure<CreateClaimsMappingPolicyRequest>(prefix + "constant Value is required.");
          // Constant values are data: preserve intentional whitespace and punctuation.
          claims.Add(claim with { Source = null, Id = null, JwtClaimType = jwt, SamlClaimType = saml });
          break;
        default:
          return Result.Failure<CreateClaimsMappingPolicyRequest>(prefix + "select a supported value mode.");
      }
    }
    return this with { DisplayName = DisplayName.Trim(), ClaimsSchema = claims.AsReadOnly() };
  }

  private static string? Optional(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}

public static class ClaimsMappingClaimSources
{
  public static IReadOnlyList<string> Sources { get; } = Array.AsReadOnly<string>(["user", "application", "resource", "audience", "company"]);
  private static readonly IReadOnlyList<string> UserProperties = Array.AsReadOnly<string>([
    "surname", "givenname", "displayname", "objectid", "mail", "userprincipalname", "department",
    "onpremisessamaccountname", "companyname", "streetaddress", "postalcode", "preferredlanguage",
    "mobilephone", "officelocation", "usertype", "telephonenumber"]);
  private static readonly IReadOnlyList<string> ApplicationProperties = Array.AsReadOnly<string>(["displayname", "objectid", "tags"]);
  private static readonly IReadOnlyList<string> CompanyProperties = Array.AsReadOnly<string>(["tenantcountry"]);

  public static bool IsSupported(string? source) => source is not null && Sources.Contains(source, StringComparer.Ordinal);

  public static IReadOnlyList<string> PropertiesFor(string? source) => source switch
  {
    "user" => UserProperties,
    "application" or "resource" or "audience" => ApplicationProperties,
    "company" => CompanyProperties,
    _ => []
  };

  public static string Label(string source) => source switch
  {
    "user" => "User",
    "application" => "Application (client)",
    "resource" => "Resource service principal",
    "audience" => "Audience service principal",
    "company" => "Company (tenant)",
    _ => source
  };
}
