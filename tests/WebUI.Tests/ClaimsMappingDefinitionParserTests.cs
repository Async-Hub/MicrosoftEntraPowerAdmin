using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.ClaimsMappingPolicies;
using System.Text.Json;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingDefinitionParserTests
{
  public const string Definition = """
    {"ClaimsMappingPolicy":{"Version":1,"IncludeBasicClaimSet":"true",
    "ClaimsSchema":[{"Source":"user","ID":"department","JwtClaimType":"department","FutureField":{"nested":42}},
    {"Source":"transformation","ID":"joined","SamlClaimType":"urn:joined","Value":"fixed","TransformationId":"Join"}],
    "ClaimsTransformation":[{"ID":"Join","TransformationMethod":"Join",
    "InputClaims":[{"ClaimTypeReferenceId":"department","TransformationClaimType":"string1"}],
    "InputParameters":[{"ID":"separator","DataType":"string","Value":"."}],
    "OutputClaims":[{"ClaimTypeReferenceId":"joined","TransformationClaimType":"outputClaim"}]}],
    "UnknownPolicyProperty":[1,2]},"FutureRoot":{"enabled":true}}
    """;

  [Fact]
  public void Parses_known_policy_claim_and_transformation_fields()
  {
    var result = ClaimsMappingDefinitionParser.Parse(Definition);
    Assert.Empty(result.Warnings);
    var document = Assert.IsType<ClaimsMappingDocument>(result.Document);
    Assert.Equal(1, document.Version);
    Assert.True(document.IncludeBasicClaimSet);
    Assert.Equal(2, document.ClaimsSchema?.Count);
    Assert.Equal("department", document.ClaimsSchema?[0].JwtClaimType);
    Assert.Equal("urn:joined", document.ClaimsSchema?[1].SamlClaimType);
    Assert.Equal("fixed", document.ClaimsSchema?[1].Value);
    var transformation = Assert.Single(document.ClaimsTransformation!);
    Assert.Equal("Join", transformation.TransformationMethod);
    Assert.Equal("department", Assert.Single(transformation.InputClaims!).ClaimTypeReferenceId);
    Assert.Equal(".", Assert.Single(transformation.InputParameters!).Value);
    Assert.Equal("outputClaim", Assert.Single(transformation.OutputClaims!).TransformationClaimType);
  }

  [Fact]
  public void Original_and_formatted_definition_preserve_unknown_properties()
  {
    var result = ClaimsMappingDefinitionParser.Parse(Definition);

    Assert.Equal(Definition, result.Raw);
    using var original = JsonDocument.Parse(Definition);
    using var formatted = JsonDocument.Parse(result.Formatted);
    Assert.True(JsonElement.DeepEquals(original.RootElement, formatted.RootElement));
  }

  [Theory]
  [InlineData("")]
  [InlineData("  ")]
  [InlineData("{broken")]
  [InlineData("\"JSON inside another JSON string\"")]
  [InlineData("[]")]
  [InlineData("null")]
  [InlineData("{\"OtherPolicy\":{}}")]
  [InlineData("{\"ClaimsMappingPolicy\":[]}")]
  public void Empty_malformed_and_unknown_structures_remain_inspectable(string raw)
  {
    var result = ClaimsMappingDefinitionParser.Parse(raw);
    Assert.Equal(raw, result.Raw);
    Assert.Null(result.Document);
    Assert.NotEmpty(result.Warnings);
    if (raw == "{broken") Assert.Equal(raw, result.Formatted);
  }

  [Fact]
  public void Unexpected_known_fields_do_not_discard_other_understood_entries()
  {
    var result = ClaimsMappingDefinitionParser.Parse("""
      {"ClaimsMappingPolicy":{"Version":{},"IncludeBasicClaimSet":[],
      "ClaimsSchema":[false,{"Source":"user","ID":{},"Value":42}],
      "ClaimsTransformation":[{"ID":"transform","InputClaims":{},"OutputClaims":[null,{}]}]}}
      """);
    var document = Assert.IsType<ClaimsMappingDocument>(result.Document);
    Assert.Null(document.Version);
    Assert.Null(document.IncludeBasicClaimSet);
    var entry = Assert.Single(document.ClaimsSchema!);
    Assert.Equal("user", entry.Source);
    Assert.Null(entry.Id);
    Assert.Equal("42", entry.Value);
    var transform = Assert.Single(document.ClaimsTransformation!);
    Assert.Null(transform.InputClaims);
    Assert.Single(transform.OutputClaims!);
    Assert.NotEmpty(result.Warnings);
  }

  [Theory]
  [InlineData("true", true)]
  [InlineData("false", false)]
  [InlineData("\"false\"", false)]
  public void Include_basic_claim_set_accepts_Boolean_and_string_Boolean_values(string jsonBoolean, bool expected)
  {
    var result = ClaimsMappingDefinitionParser.Parse(
      "{\"ClaimsMappingPolicy\":{\"IncludeBasicClaimSet\":" + jsonBoolean + "}}");
    Assert.Equal(expected, result.Document?.IncludeBasicClaimSet);
  }

  [Theory]
  [InlineData("ClaimsTransformation")]
  [InlineData("ClaimsTransformations")]
  public void Singular_and_plural_transformation_names_preserve_an_explicit_empty_collection(string property)
  {
    var result = ClaimsMappingDefinitionParser.Parse("{\"ClaimsMappingPolicy\":{\"" + property + "\":[]}}");

    var document = Assert.IsType<ClaimsMappingDocument>(result.Document);
    Assert.Empty(document.ClaimsTransformation!);
    Assert.Null(document.ClaimsSchema);
  }

  [Theory]
  [InlineData("broken")]
  [InlineData(null)]
  public void Summary_does_not_report_totals_for_partial_or_missing_definitions(string? raw)
  {
    var policy = new ClaimsMappingPolicyListItem(Guid.NewGuid(), "Policy", null,
      raw is null ? [] : [ClaimsMappingDefinitionParser.Parse(Definition), ClaimsMappingDefinitionParser.Parse(raw)]);

    Assert.Null(policy.ClaimsSchemaCount);
    Assert.Null(policy.ClaimsTransformationCount);
  }
}
