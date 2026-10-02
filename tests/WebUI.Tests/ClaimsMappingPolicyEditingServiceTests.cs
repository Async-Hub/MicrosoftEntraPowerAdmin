using System.Net;
using System.Text;
using System.Text.Json;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using CSharpFunctionalExtensions;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingPolicyEditingServiceTests
{
  [Fact]
  public async Task Begin_edit_reads_fresh_policy_and_assignments()
  {
    using var harness = new Harness();
    harness.Name = "Latest remote name";
    var context = await harness.BeginAsync();
    Assert.Equal("Latest remote name", context.Original.Policy.DisplayName);
    Assert.Equal(harness.Id, context.Original.Policy.ObjectId);
    Assert.Equal(harness.Tenants.TenantA, context.TenantId);
    Assert.Equal(new[] { harness.Path, harness.Path + "/appliesTo" }, harness.Requests.Select(request => request.Path));
  }

  [Theory]
  [InlineData("{broken")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"ClaimsTransformation\":[{\"ID\":\"keep-me\"}]}}")]
  [InlineData(ClaimsMappingPolicyEditingTests.Supported)]
  public async Task Rename_PATCH_omits_definition_and_preserves_ID_assignments_and_raw_content(string definition)
  {
    using var harness = new Harness { Definition = definition, Assigned = true };
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, new("Renamed"), TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    var patch = Assert.Single(harness.Requests, request => request.Method == HttpMethod.Patch);
    using var body = JsonDocument.Parse(patch.Body);
    Assert.Equal("Renamed", body.RootElement.GetProperty("displayName").GetString());
    Assert.False(body.RootElement.TryGetProperty("definition", out _));
    Assert.False(body.RootElement.TryGetProperty("isOrganizationDefault", out _));
    Assert.Equal(harness.Id, result.Value.Policy.ObjectId);
    Assert.Equal(definition, Assert.Single(result.Value.Policy.Definitions).Raw);
    Assert.Equal(harness.PrincipalId, Assert.Single(result.Value.Assignments.Value.ServicePrincipals).ObjectId);
    Assert.All(harness.Requests, request => Assert.True(request.Method == HttpMethod.Patch || request.Method == HttpMethod.Get));
    Assert.Equal(new[] { HttpMethod.Patch, HttpMethod.Get, HttpMethod.Get }, harness.Requests.Select(request => request.Method));
  }

  [Fact]
  public async Task Definition_update_rechecks_remote_state_uses_PATCH_and_reloads_authoritative_response_after_204()
  {
    using var harness = new Harness { NameAfterPatch = "Graph authoritative name" };
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, Changed(context), TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    var patch = Assert.Single(harness.Requests, request => request.Method == HttpMethod.Patch);
    Assert.Equal(harness.Path, patch.Path);
    using var body = JsonDocument.Parse(patch.Body);
    Assert.False(body.RootElement.TryGetProperty("displayName", out _));
    using var inner = JsonDocument.Parse(Assert.Single(body.RootElement.GetProperty("definition").EnumerateArray()).GetString()!);
    Assert.Equal("false", inner.RootElement.GetProperty("ClaimsMappingPolicy").GetProperty("IncludeBasicClaimSet").GetString());
    Assert.Equal("Graph authoritative name", result.Value.Policy.DisplayName);
    Assert.False(result.Value.Policy.Definitions[0].Document!.IncludeBasicClaimSet);
    Assert.Equal(new[] { HttpMethod.Get, HttpMethod.Get, HttpMethod.Patch, HttpMethod.Get, HttpMethod.Get }, harness.Requests.Select(request => request.Method));
  }

  [Fact]
  public async Task Unchanged_draft_performs_no_Graph_operation_or_client_acquisition()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var clients = harness.ClientCount;
    var result = await harness.Service.UpdateAsync(context, new(context.Original.Policy.DisplayName,
      ClaimsMappingDefinitionEditing.Inspect(context.Original.Policy).Definition.Value), TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Same(context.Original, result.Value);
    Assert.Empty(harness.Requests);
    Assert.Equal(clients, harness.ClientCount);
  }

  [Fact]
  public async Task Remote_definition_change_blocks_overwrite_and_preserves_latest_content()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Definition = "{\"future\":\"external change\"}";
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, Changed(context), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.StaleState, result.Error.Type);
    Assert.Contains("Reload", result.Error.Message);
    Assert.Single(harness.Requests);
    Assert.Equal(HttpMethod.Get, harness.Requests[0].Method);
    Assert.Contains("external change", harness.Definition);
  }

  [Fact]
  public async Task Unsafe_definition_replacement_is_rejected_before_any_Graph_operation()
  {
    using var harness = new Harness { Definition = "{broken" };
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var count = harness.ClientCount;
    var result = await harness.Service.UpdateAsync(context, new("Renamed", new("Renamed", true, [])), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.InvalidInput, result.Error.Type);
    Assert.Empty(harness.Requests);
    Assert.Equal(count, harness.ClientCount);
    Assert.Equal("{broken", harness.Definition);
  }

  [Fact]
  public async Task Failed_reload_after_successful_PATCH_reports_update_and_requires_refresh()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.BeforeSend = (request, _) =>
    {
      if (request.Method == HttpMethod.Patch) harness.ReadStatus = HttpStatusCode.Forbidden;
      return Task.CompletedTask;
    };
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, new("Renamed"), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.Forbidden, result.Error.Type);
    Assert.Contains("policy was updated", result.Error.Message);
    Assert.Contains("Reload", result.Error.Message);
    Assert.Equal("Renamed", harness.Name);
    Assert.Single(harness.Requests, request => request.Method == HttpMethod.Patch);
  }

  [Fact]
  public async Task Failed_assignment_discovery_prevents_definition_PATCH()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.BeforeSend = (request, _) =>
    {
      if (request.RequestUri!.AbsolutePath.EndsWith("/appliesTo"))
        throw new Microsoft.Kiota.Abstractions.ApiException { ResponseStatusCode = 403 };
      return Task.CompletedTask;
    };
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, Changed(context), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.Forbidden, result.Error.Type);
    Assert.DoesNotContain(harness.Requests, request => request.Method == HttpMethod.Patch);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Definition_update_of_assigned_policy_requires_current_assignment_confirmation(bool confirmed)
  {
    using var harness = new Harness { Assigned = true };
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, Changed(context) with
    {
      ConfirmedAssignmentIds = confirmed ? [harness.PrincipalId] : null
    }, TestContext.Current.CancellationToken);
    Assert.Equal(confirmed, result.IsSuccess);
    if (!confirmed)
    {
      Assert.Equal(GraphOperationErrorType.ConfirmationRequired, result.Error.Type);
      Assert.DoesNotContain(harness.Requests, request => request.Method == HttpMethod.Patch);
    }
    else Assert.Single(result.Value.Assignments.Value.ServicePrincipals);
  }

  [Fact]
  public async Task Newly_added_assignment_blocks_unconfirmed_definition_update()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Assigned = true;
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, Changed(context), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.ConfirmationRequired, result.Error.Type);
    Assert.DoesNotContain(harness.Requests, request => request.Method == HttpMethod.Patch);
  }

  [Theory]
  [InlineData(400, GraphOperationErrorType.InvalidInput)]
  [InlineData(403, GraphOperationErrorType.Forbidden)]
  [InlineData(404, GraphOperationErrorType.NotFound)]
  [InlineData(429, GraphOperationErrorType.Throttled)]
  [InlineData(503, GraphOperationErrorType.GraphFailure)]
  public async Task PATCH_errors_use_safe_messages_and_preserve_diagnostics(int status, GraphOperationErrorType type)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.PatchStatus = (HttpStatusCode)status;
    var result = await harness.Service.UpdateAsync(context, new("Renamed"), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(type, result.Error.Type);
    Assert.Equal(status, result.Error.HttpStatus);
    Assert.Equal("test-error", result.Error.Code);
    Assert.Equal("edit-request", result.Error.RequestId);
    Assert.DoesNotContain("sensitive", result.Error.Message);
    if (status == 403)
    {
      Assert.Contains("Policy.ReadWrite.ApplicationConfiguration", result.Error.Message);
      Assert.DoesNotContain("Application.ReadWrite.All", result.Error.Message);
    }
  }

  [Fact]
  public async Task Deleted_policy_during_fresh_state_check_returns_not_found_without_recreation()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Deleted = true;
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, Changed(context), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.NotFound, result.Error.Type);
    Assert.All(harness.Requests, request => Assert.Equal(HttpMethod.Get, request.Method));
  }

  [Fact]
  public async Task Consent_failure_mentions_existing_policy_write_permission()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.BeforeClient = () => throw new MsalUiRequiredException("consent_required", "sensitive error");
    var result = await harness.Service.UpdateAsync(context, new("Renamed"), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.ConsentRequired, result.Error.Type);
    Assert.Contains("Policy.ReadWrite.ApplicationConfiguration", result.Error.Message);
    Assert.DoesNotContain("sensitive", result.Error.Message);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Tenant_switch_invalidates_draft_even_after_switching_back(bool switchBack)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    if (switchBack) harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantA);
    harness.Requests.Clear();
    var count = harness.ClientCount;
    var result = await harness.Service.UpdateAsync(context, new("Renamed"), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Empty(harness.Requests);
    Assert.Equal(count, harness.ClientCount);
  }

  [Fact]
  public async Task Tenant_switch_during_client_acquisition_prevents_even_name_only_PATCH()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.BeforeClient = () => harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, new("Renamed"), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Empty(harness.Requests);
  }

  [Fact]
  public async Task Tenant_switch_during_remote_read_prevents_definition_PATCH()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.BeforeSend = (_, _) =>
    {
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      return Task.CompletedTask;
    };
    harness.Requests.Clear();
    var result = await harness.Service.UpdateAsync(context, Changed(context), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Equal(HttpMethod.Get, Assert.Single(harness.Requests).Method);
  }

  [Fact]
  public async Task Caller_cancellation_is_propagated_to_Graph()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    using var cancellation = new CancellationTokenSource();
    harness.BeforeSend = async (_, token) =>
    {
      cancellation.Cancel();
      await Task.Delay(Timeout.Infinite, token);
    };
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => harness.Service.UpdateAsync(context, new("Renamed"), cancellation.Token));
  }

  [Fact]
  public async Task Late_PATCH_response_does_not_reload_old_policy_ID_in_new_tenant()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    var pending = new TaskCompletionSource();
    harness.BeforeSend = (_, _) => pending.Task;
    harness.Requests.Clear();
    var save = harness.Service.UpdateAsync(context, new("Renamed"), TestContext.Current.CancellationToken);
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    pending.SetResult();
    var result = await save;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Equal(HttpMethod.Patch, Assert.Single(harness.Requests).Method);
  }

  private static UpdateClaimsMappingPolicyRequest Changed(ClaimsMappingPolicyEditContext context) =>
    new(context.Original.Policy.DisplayName, ClaimsMappingDefinitionEditing.Inspect(context.Original.Policy).Definition.Value with { IncludeBasicClaimSet = false });

  private sealed class Harness : IEntraGraphClientFactory, IDisposable
  {
    private readonly DirectoryReadOperation _operation;
    public GraphTestContext Tenants { get; } = new();
    public Guid Id { get; } = Guid.NewGuid();
    public Guid PrincipalId { get; } = Guid.NewGuid();
    public string Path => $"/v1.0/policies/claimsMappingPolicies/{Id:D}";
    public string Name { get; set; } = "Policy";
    public string Definition { get; set; } = ClaimsMappingPolicyEditingTests.Supported;
    public string? NameAfterPatch { get; set; }
    public bool Assigned { get; set; }
    public bool Deleted { get; set; }
    public HttpStatusCode PatchStatus { get; set; } = HttpStatusCode.NoContent;
    public HttpStatusCode ReadStatus { get; set; } = HttpStatusCode.OK;
    public Action? BeforeClient { get; set; }
    public Func<HttpRequestMessage, CancellationToken, Task>? BeforeSend { get; set; }
    public List<(HttpMethod Method, string Path, string Body)> Requests { get; } = [];
    public ClaimsMappingPolicyService Service { get; }
    public int ClientCount { get; private set; }

    public Harness()
    {
      Tenants.Tenants.SelectTenant(Tenants.TenantA);
      _operation = new(Tenants.Tenants, this, NullLogger<DirectoryReadOperation>.Instance);
      Service = new(_operation, Tenants.Tenants, this, NullLogger<ClaimsMappingPolicyService>.Instance);
    }

    public async Task<ClaimsMappingPolicyEditContext> BeginAsync()
    {
      var result = await Service.BeginEditAsync(Id, TestContext.Current.CancellationToken);
      Assert.True(result.IsSuccess);
      return result.Value;
    }

    public Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(CancellationToken cancellationToken = default)
    {
      ClientCount++;
      BeforeClient?.Invoke();
      return Task.FromResult(Result.Success<GraphServiceClient, GraphOperationError>(new GraphServiceClient(
        new HttpClient(new Transport(SendAsync)), new AnonymousAuthenticationProvider())));
    }

    private async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
    {
      var body = request.Content is null ? "" : await request.Content.ReadAsStringAsync(token);
      var path = request.RequestUri!.AbsolutePath;
      Requests.Add((request.Method, path, body));
      Assert.StartsWith(Path, path);
      if (BeforeSend is { } before) await before(request, token);
      if (Deleted || (request.Method == HttpMethod.Patch && PatchStatus != HttpStatusCode.NoContent) ||
          (request.Method == HttpMethod.Get && ReadStatus != HttpStatusCode.OK))
        return Json(new { error = new { code = "test-error", message = "sensitive Graph error", innerError = new Dictionary<string, string> { ["request-id"] = "edit-request" } } },
          Deleted ? HttpStatusCode.NotFound : request.Method == HttpMethod.Patch ? PatchStatus : ReadStatus);
      if (request.Method == HttpMethod.Patch)
      {
        using var json = JsonDocument.Parse(body);
        if (json.RootElement.TryGetProperty("displayName", out var name)) Name = name.GetString()!;
        if (json.RootElement.TryGetProperty("definition", out var definition)) Definition = definition[0].GetString()!;
        if (NameAfterPatch is { } authoritative) Name = authoritative;
        return new(HttpStatusCode.NoContent);
      }
      Assert.Equal(HttpMethod.Get, request.Method);
      if (path.EndsWith("/appliesTo"))
        return Json(new
        {
          value = Assigned ? new object[] { new Dictionary<string, object>
        {
          ["@odata.type"] = "#microsoft.graph.servicePrincipal", ["id"] = PrincipalId,
          ["appId"] = Guid.NewGuid(), ["displayName"] = "Assigned Enterprise"
        } } : []
        });
      return Json(new { id = Id, displayName = Name, definition = new[] { Definition }, isOrganizationDefault = false });
    }

    private static HttpResponseMessage Json(object value, HttpStatusCode status = HttpStatusCode.OK) => new(status)
    {
      Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json")
    };
    public void Dispose() => _operation.Dispose();
  }

  private sealed class Transport(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) : HttpMessageHandler
  {
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => send(request, cancellationToken);
  }
}
