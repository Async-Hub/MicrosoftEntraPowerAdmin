using System.Text.Json;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingPolicyEditingTests
{
  internal const string Supported = """
    {"ClaimsMappingPolicy":{"Version":1,"IncludeBasicClaimSet":"true","ClaimsSchema":[
      {"Source":"user","ID":"department","JwtClaimType":"department"},
      {"Value":"  West \"region\"  ","SamlClaimType":"https://example.test/region"}]}}
    """;

  [Fact]
  public void Supported_definition_round_trips_all_values_and_both_claim_modes()
  {
    var editing = ClaimsMappingDefinitionEditing.Inspect(Policy(Supported));
    Assert.Equal(ClaimsMappingDefinitionEditability.FullySupported, editing.Classification);
    Assert.True(editing.CanEditDefinition);
    Assert.Equal(ClaimValueMode.DirectoryAttribute, editing.Definition.Value.ClaimsSchema[0].Mode);
    Assert.Equal(ClaimValueMode.ConstantValue, editing.Definition.Value.ClaimsSchema[1].Mode);
    using var original = JsonDocument.Parse(Supported);
    using var serialized = JsonDocument.Parse(ClaimsMappingPolicyDefinitionSerializer.Serialize(editing.Definition.Value));
    Assert.True(JsonElement.DeepEquals(original.RootElement, serialized.RootElement));
  }

  [Theory]
  [InlineData("\"ClaimsTransformation\":[]")]
  [InlineData("\"ClaimsTransformations\":[]")]
  [InlineData("\"FutureMicrosoftProperty\":{\"Something\":true}")]
  [InlineData("\"Version\":2")]
  public void Unsupported_policy_features_allow_rename_but_never_definition_replacement(string extra)
  {
    var raw = "{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[]," + extra + "}}";
    AssertProtected(raw);
  }

  [Theory]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":2,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[]}}")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[]}}")]
  [InlineData("{\"FutureRoot\":true,\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[]}}")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[{\"Source\":\"user\",\"ID\":\"mail\",\"JwtClaimType\":\"email\",\"Future\":true}]}}")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[{\"Source\":\"user\",\"ID\":\"mail\",\"JwtClaimType\":\"email\",\"Value\":\"discarded\"}]}}")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":true,\"ClaimsSchema\":[]}}")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[{\"Source\":\"user\",\"ID\":\"mail\",\"ID\":\"department\",\"JwtClaimType\":\"email\"}]}}")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[{\"Source\":\"user\",\"ID\":\" mail \",\"JwtClaimType\":\"email\"}]}}")]
  public void Unknown_shapes_types_duplicates_and_normalization_loss_disable_definition_editing(string raw) => AssertProtected(raw);

  [Fact]
  public void Multiple_or_missing_definition_strings_are_not_rewritten()
  {
    var multiple = Policy(Supported) with { Definitions = [ClaimsMappingDefinitionParser.Parse(Supported), ClaimsMappingDefinitionParser.Parse(Supported)] };
    Assert.False(ClaimsMappingDefinitionEditing.Inspect(multiple).CanEditDefinition);
    Assert.False(ClaimsMappingDefinitionEditing.Inspect(multiple with { Definitions = [] }).CanEditDefinition);
  }

  [Fact]
  public void Malformed_JSON_remains_raw_and_can_be_renamed_without_repair()
  {
    var policy = Policy("{broken");
    var editing = ClaimsMappingDefinitionEditing.Inspect(policy);
    Assert.Equal(ClaimsMappingDefinitionEditability.MalformedDefinition, editing.Classification);
    Assert.False(editing.CanEditDefinition);
    Assert.Equal("{broken", policy.Definitions[0].Raw);
    var patch = new UpdateClaimsMappingPolicyRequest("Renamed").CreatePatch(policy);
    Assert.True(patch.IsSuccess);
    Assert.Null(patch.Value.Value.Definition);
  }

  [Fact]
  public void Unchanged_and_reverted_drafts_do_not_produce_an_update()
  {
    var policy = Policy(Supported);
    var definition = ClaimsMappingDefinitionEditing.Inspect(policy).Definition.Value;
    Assert.True(new UpdateClaimsMappingPolicyRequest(policy.DisplayName).CreatePatch(policy).Value.HasNoValue);
    Assert.True(new UpdateClaimsMappingPolicyRequest(policy.DisplayName, definition).CreatePatch(policy).Value.HasNoValue);
    var draft = ClaimsMappingPolicyDraft.From(definition);
    draft.IncludeBasicClaimSet = false;
    draft.IncludeBasicClaimSet = true;
    Assert.True(new UpdateClaimsMappingPolicyRequest(policy.DisplayName, draft.ToRequest()).CreatePatch(policy).Value.HasNoValue);
  }

  [Fact]
  public void Existing_name_whitespace_is_preserved_until_the_administrator_changes_it()
  {
    var policy = Policy(Supported) with { DisplayName = "  Policy  " };
    var unchanged = new UpdateClaimsMappingPolicyRequest(policy.DisplayName,
      ClaimsMappingDefinitionEditing.Inspect(policy).Definition.Value).CreatePatch(policy);
    Assert.True(unchanged.IsSuccess);
    Assert.True(unchanged.Value.HasNoValue);
    var renamed = new UpdateClaimsMappingPolicyRequest("Renamed").CreatePatch(policy);
    Assert.Equal("Renamed", renamed.Value.Value.DisplayName);
    Assert.Null(renamed.Value.Value.Definition);
  }

  [Fact]
  public void Definition_change_sends_version_one_string_boolean_and_keeps_unchanged_name_absent()
  {
    var policy = Policy(Supported);
    var draft = ClaimsMappingPolicyDraft.From(ClaimsMappingDefinitionEditing.Inspect(policy).Definition.Value);
    draft.IncludeBasicClaimSet = false;
    draft.Claims[0].Source = "application";
    draft.Claims[0].Id = "displayname";
    draft.Claims.RemoveAt(1);
    draft.Claims.Add(new() { Mode = ClaimValueMode.ConstantValue, Value = "constant", JwtClaimType = "new-claim" });
    var patch = new UpdateClaimsMappingPolicyRequest(policy.DisplayName, draft.ToRequest()).CreatePatch(policy);
    Assert.True(patch.IsSuccess);
    Assert.Null(patch.Value.Value.DisplayName);
    using var json = JsonDocument.Parse(Assert.Single(patch.Value.Value.Definition!));
    var inner = json.RootElement.GetProperty("ClaimsMappingPolicy");
    Assert.Equal(1, inner.GetProperty("Version").GetInt32());
    Assert.Equal("false", inner.GetProperty("IncludeBasicClaimSet").GetString());
    Assert.Equal("application", inner.GetProperty("ClaimsSchema")[0].GetProperty("Source").GetString());
    Assert.Equal("displayname", inner.GetProperty("ClaimsSchema")[0].GetProperty("ID").GetString());
    Assert.Equal("constant", inner.GetProperty("ClaimsSchema")[1].GetProperty("Value").GetString());
    Assert.False(inner.GetProperty("ClaimsSchema")[1].TryGetProperty("Source", out _));
    Assert.Equal(Supported, policy.Definitions[0].Raw);
  }

  [Fact]
  public void Editing_reuses_creation_validation_including_name_placeholder_rows_and_claim_limit()
  {
    var policy = Policy(Supported);
    var request = ClaimsMappingDefinitionEditing.Inspect(policy).Definition.Value;
    Assert.Contains("Display Name", new UpdateClaimsMappingPolicyRequest(" ").CreatePatch(policy).Error);
    Assert.Contains("50", new UpdateClaimsMappingPolicyRequest("Policy", request with
    {
      ClaimsSchema = Enumerable.Repeat(request.ClaimsSchema[0], 51).ToArray()
    }).CreatePatch(policy).Error);
    var draft = ClaimsMappingPolicyDraft.From(request);
    draft.Claims.Add(new());
    Assert.Contains("JWT or SAML", new UpdateClaimsMappingPolicyRequest("Policy", draft.ToRequest()).CreatePatch(policy).Error);
  }

  private static void AssertProtected(string raw)
  {
    var policy = Policy(raw);
    var editing = ClaimsMappingDefinitionEditing.Inspect(policy);
    Assert.Equal(ClaimsMappingDefinitionEditability.UnsupportedDefinition, editing.Classification);
    Assert.False(editing.CanEditDefinition);
    var renamed = new UpdateClaimsMappingPolicyRequest("Renamed").CreatePatch(policy);
    Assert.True(renamed.IsSuccess);
    Assert.Equal("Renamed", renamed.Value.Value.DisplayName);
    Assert.Null(renamed.Value.Value.Definition);
    Assert.Equal(raw, policy.Definitions[0].Raw);
    var replacement = new UpdateClaimsMappingPolicyRequest("Renamed", new("Renamed", true, [])).CreatePatch(policy);
    Assert.True(replacement.IsFailure);
    Assert.Contains("disabled", replacement.Error);
  }

  internal static ClaimsMappingPolicyListItem Policy(string raw) => new(Guid.NewGuid(), "Policy", false,
    [ClaimsMappingDefinitionParser.Parse(raw)]);
}
