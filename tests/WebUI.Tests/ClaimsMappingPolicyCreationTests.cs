using System.Text.Json;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingPolicyCreationTests
{
  [Theory]
  [InlineData(true, "true")]
  [InlineData(false, "false")]
  public void Basic_claim_has_documented_names_version_and_string_boolean(bool includeBasic, string expected)
  {
    using var document = Serialize([DirectoryClaim()], includeBasic);
    var policy = document.RootElement.GetProperty("ClaimsMappingPolicy");
    Assert.Equal(1, policy.GetProperty("Version").GetInt32());
    Assert.Equal(expected, policy.GetProperty("IncludeBasicClaimSet").GetString());
    var claim = Assert.Single(policy.GetProperty("ClaimsSchema").EnumerateArray());
    Assert.Equal("user", claim.GetProperty("Source").GetString());
    Assert.Equal("department", claim.GetProperty("ID").GetString());
    Assert.Equal("department", claim.GetProperty("JwtClaimType").GetString());
    Assert.False(claim.TryGetProperty("SamlClaimType", out _));
    Assert.False(claim.TryGetProperty("Value", out _));
    Assert.False(policy.TryGetProperty("ClaimsTransformation", out _));
    Assert.False(policy.TryGetProperty("ClaimsTransformations", out _));
  }

  [Fact]
  public void Constant_claim_omits_directory_properties_and_preserves_value()
  {
    const string value = "  A \"quoted\" value \\ with newline\n  ";
    var input = new CreateClaimsSchemaEntry(ClaimValueMode.ConstantValue, "user", "department", "constant", " ", value);
    using var document = Serialize([input]);
    var claim = document.RootElement.GetProperty("ClaimsMappingPolicy").GetProperty("ClaimsSchema")[0];
    Assert.Equal(value, claim.GetProperty("Value").GetString());
    Assert.False(claim.TryGetProperty("Source", out _));
    Assert.False(claim.TryGetProperty("ID", out _));
    Assert.False(claim.TryGetProperty("SamlClaimType", out _));
  }

  [Fact]
  public void Saml_only_claim_keeps_URI_and_omits_empty_JWT_and_irrelevant_value()
  {
    const string uri = "https://example.test/claims/Department?name=CaseSensitive";
    using var document = Serialize([DirectoryClaim() with { JwtClaimType = " ", SamlClaimType = uri, Value = "unused" }]);
    var claim = document.RootElement.GetProperty("ClaimsMappingPolicy").GetProperty("ClaimsSchema")[0];
    Assert.Equal(uri, claim.GetProperty("SamlClaimType").GetString());
    Assert.False(claim.TryGetProperty("JwtClaimType", out _));
    Assert.False(claim.TryGetProperty("Value", out _));
  }

  [Fact]
  public void Multiple_claims_preserve_order_and_allow_both_output_types()
  {
    using var document = Serialize([
      DirectoryClaim() with { SamlClaimType = "department-saml" },
      new(ClaimValueMode.ConstantValue, null, null, "region", null, "West")]);
    var claims = document.RootElement.GetProperty("ClaimsMappingPolicy").GetProperty("ClaimsSchema");
    Assert.Equal(2, claims.GetArrayLength());
    Assert.Equal("department-saml", claims[0].GetProperty("SamlClaimType").GetString());
    Assert.Equal("department", claims[0].GetProperty("JwtClaimType").GetString());
    Assert.Equal("West", claims[1].GetProperty("Value").GetString());
  }

  [Theory]
  [InlineData(null)]
  [InlineData("")]
  [InlineData(" \t ")]
  public void Empty_policy_names_fail_validation(string? name)
  {
    var result = new CreateClaimsMappingPolicyRequest(name!, true, []).ValidateAndNormalize();
    Assert.True(result.IsFailure);
    Assert.Equal("Display Name is required.", result.Error);
  }

  [Theory]
  [InlineData(ClaimValueMode.DirectoryAttribute, "transformation", "department", "department", null, null, "supported directory source")]
  [InlineData(ClaimValueMode.DirectoryAttribute, null, "department", "department", null, null, "supported directory source")]
  [InlineData(ClaimValueMode.DirectoryAttribute, "user", " ", "department", null, null, "attribute ID")]
  [InlineData(ClaimValueMode.ConstantValue, null, null, "constant", null, " ", "constant Value")]
  [InlineData(ClaimValueMode.DirectoryAttribute, "user", "department", " ", null, null, "JWT or SAML")]
  [InlineData(ClaimValueMode.ConstantValue, null, null, null, " ", "value", "JWT or SAML")]
  [InlineData((ClaimValueMode)99, null, null, "type", null, "value", "supported value mode")]
  public void Incomplete_or_unsupported_claims_are_rejected(ClaimValueMode mode, string? source, string? id,
    string? jwt, string? saml, string? value, string error)
  {
    var result = new CreateClaimsMappingPolicyRequest("Policy", true, [new(mode, source, id, jwt, saml, value)]).ValidateAndNormalize();
    Assert.True(result.IsFailure);
    Assert.Contains(error, result.Error);
    Assert.Contains("Claim 1", result.Error);
  }

  [Theory]
  [InlineData(50, true)]
  [InlineData(51, false)]
  public void Enforces_claim_limit(int count, bool valid)
  {
    var result = new CreateClaimsMappingPolicyRequest("Policy", true,
      Enumerable.Repeat(DirectoryClaim(), count).ToArray()).ValidateAndNormalize();
    Assert.Equal(valid, result.IsSuccess);
    if (!valid) Assert.Contains("50", result.Error);
  }

  [Theory]
  [InlineData("user")]
  [InlineData("application")]
  [InlineData("resource")]
  [InlineData("audience")]
  [InlineData("company")]
  public void Supported_sources_allow_custom_property_IDs_and_trim_input(string source)
  {
    var result = new CreateClaimsMappingPolicyRequest("  Policy  ", true,
      [DirectoryClaim() with { Source = $" {source} ", Id = " customproperty ", JwtClaimType = " output " }]).ValidateAndNormalize();
    Assert.True(result.IsSuccess);
    Assert.Equal("Policy", result.Value.DisplayName);
    var claim = Assert.Single(result.Value.ClaimsSchema);
    Assert.Equal(source, claim.Source);
    Assert.Equal("customproperty", claim.Id);
    Assert.Equal("output", claim.JwtClaimType);
  }

  [Fact]
  public void Zero_claims_allow_a_policy_containing_only_basic_claim_set_configuration()
  {
    using var document = Serialize([]);
    Assert.Equal(0, document.RootElement.GetProperty("ClaimsMappingPolicy").GetProperty("ClaimsSchema").GetArrayLength());
  }

  private static CreateClaimsSchemaEntry DirectoryClaim() =>
    new(ClaimValueMode.DirectoryAttribute, "user", "department", "department", null, null);

  private static JsonDocument Serialize(IReadOnlyList<CreateClaimsSchemaEntry> claims, bool includeBasic = true)
  {
    var validated = new CreateClaimsMappingPolicyRequest("Policy", includeBasic, claims).ValidateAndNormalize();
    Assert.True(validated.IsSuccess);
    return JsonDocument.Parse(ClaimsMappingPolicyDefinitionSerializer.Serialize(validated.Value));
  }
}
