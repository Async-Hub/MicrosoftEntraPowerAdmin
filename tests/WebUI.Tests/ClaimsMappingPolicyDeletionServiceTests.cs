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

public sealed class ClaimsMappingPolicyDeletionServiceTests
{
  [Fact]
  public async Task Begin_delete_loads_current_identity_definition_and_assignments_on_one_client()
  {
    using var harness = new Harness { Name = "Fresh name", AssignmentCount = 3 };
    var context = await harness.BeginAsync();
    Assert.Equal("Fresh name", context.CurrentDisplayName);
    Assert.Equal(harness.Id, context.Original.Policy.ObjectId);
    Assert.Equal(harness.Tenants.TenantA, context.TenantId);
    Assert.Equal(3, context.AssignmentCount);
    Assert.True(context.HasAssignments);
    Assert.Equal(1, harness.ClientCount);
    Assert.Equal(new[] { harness.Path, harness.Path + "/appliesTo" }, harness.Requests.Select(request => request.Path));
    Assert.Equal(harness.Definition, context.Original.Policy.Definitions[0].Raw);
  }

  [Theory]
  [InlineData(ClaimsMappingPolicyEditingTests.Supported)]
  [InlineData("{broken")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":99,\"Unknown\":true,\"ClaimsTransformation\":[{\"ID\":\"keep\"}]}}")]
  public async Task Zero_assignments_and_exact_confirmation_delete_only_policy_and_accept_204(string raw)
  {
    using var harness = new Harness { Definition = raw };
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "  Policy  ", TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.True(harness.Deleted);
    Assert.Equal(raw, harness.Definition);
    Assert.Equal(new[] { HttpMethod.Get, HttpMethod.Get, HttpMethod.Delete }, harness.Requests.Select(request => request.Method));
    var delete = Assert.Single(harness.Requests, request => request.Method == HttpMethod.Delete);
    Assert.Equal(harness.Path, delete.Path);
    Assert.Equal("", delete.Body);
  }

  [Theory]
  [InlineData(1)]
  [InlineData(3)]
  public async Task Assigned_policy_is_blocked_without_DELETE_or_unassignment(int count)
  {
    using var harness = new Harness { AssignmentCount = count };
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.PolicyIsAssigned, result.Error.Type);
    Assert.Contains("Unassign", result.Error.Message);
    Assert.Empty(harness.Requests);
    Assert.False(harness.Deleted);
    Assert.Equal(count, harness.AssignmentCount);
  }

  [Theory]
  [InlineData("servicePrincipal")]
  [InlineData("application")]
  [InlineData("directoryObject")]
  public async Task Newly_added_relationship_prevents_stale_delete_for_every_object_type(string objectType)
  {
    using var harness = new Harness { ObjectType = objectType };
    var context = await harness.BeginAsync();
    harness.AssignmentCount = 1;
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.PolicyIsAssigned, result.Error.Type);
    Assert.Equal(new[] { harness.Path, harness.Path + "/appliesTo" }, harness.Requests.Select(request => request.Path));
    Assert.All(harness.Requests, request => Assert.Equal(HttpMethod.Get, request.Method));
    Assert.False(harness.Deleted);
  }

  [Fact]
  public async Task Assignment_on_later_page_blocks_delete_even_when_first_page_is_empty()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.AssignmentCount = 1;
    harness.AssignmentOnNextPage = true;
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.PolicyIsAssigned, result.Error.Type);
    Assert.Equal(3, harness.Requests.Count);
    Assert.Contains("next=1", harness.Requests[^1].Query);
    Assert.All(harness.Requests, request => Assert.Equal(HttpMethod.Get, request.Method));
  }

  [Theory]
  [InlineData(null, "DELETE", true)]
  [InlineData("", "DELETE", true)]
  [InlineData("   ", "DELETE", true)]
  [InlineData("Bad\nname", "DELETE", true)]
  [InlineData("Policy", "Policy", true)]
  [InlineData("  Policy  ", " Policy ", true)]
  [InlineData("Policy", "policy", false)]
  [InlineData("Policy", "Other", false)]
  [InlineData("Policy", "", false)]
  [InlineData("Policy", null, false)]
  [InlineData("Name unavailable", "DELETE", false)]
  [InlineData("Name unavailable", "Name unavailable", true)]
  public void Confirmation_uses_exact_trimmed_name_or_explicit_fallback(string? name, string? entered, bool matches)
  {
    Assert.Equal(matches, ClaimsMappingPolicyDeletionContext.Matches(name, entered));
  }

  [Theory]
  [InlineData("")]
  [InlineData("Other")]
  [InlineData("policy")]
  public async Task Invalid_confirmation_is_rejected_before_client_acquisition(string entered)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Requests.Clear();
    var clients = harness.ClientCount;
    var result = await harness.Service.DeleteAsync(context, entered, TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.ConfirmationRequired, result.Error.Type);
    Assert.Empty(harness.Requests);
    Assert.Equal(clients, harness.ClientCount);
  }

  [Fact]
  public async Task Unknown_assignment_state_blocks_deletion_even_with_correct_name()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    context = context with
    {
      Original = context.Original with
      {
        Assignments = Result.Failure<ClaimsMappingAssignments, GraphOperationError>(new(GraphOperationErrorType.GraphFailure, "Assignments unavailable"))
      }
    };
    harness.Requests.Clear();
    Assert.Null(context.AssignmentCount);
    Assert.False(context.CanConfirm("Policy"));
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.ConfirmationRequired, result.Error.Type);
    Assert.Empty(harness.Requests);
  }

  [Theory]
  [InlineData(null)]
  [InlineData("   ")]
  [InlineData("Bad\nname")]
  public async Task Legacy_name_can_be_deleted_with_DELETE_phrase(string? name)
  {
    using var harness = new Harness { Name = name };
    var context = await harness.BeginAsync();
    Assert.Equal("DELETE", context.ConfirmationPhrase);
    var result = await harness.Service.DeleteAsync(context, "DELETE", TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.True(harness.Deleted);
  }

  [Fact]
  public async Task Renamed_policy_requires_fresh_confirmation_without_DELETE()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Name = "Renamed elsewhere";
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.StaleState, result.Error.Type);
    Assert.Equal(HttpMethod.Get, Assert.Single(harness.Requests).Method);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Not_found_on_fresh_fetch_or_DELETE_returns_stale_state(bool onDelete)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    if (onDelete) harness.DeleteStatus = HttpStatusCode.NotFound;
    else harness.Deleted = true;
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.NotFound, result.Error.Type);
    Assert.Contains("no longer exists", result.Error.Message);
    Assert.DoesNotContain(harness.Requests, request => request.Method == HttpMethod.Post || request.Method == HttpMethod.Patch);
    if (!onDelete) Assert.Equal(HttpMethod.Get, Assert.Single(harness.Requests).Method);
  }

  [Fact]
  public async Task Begin_delete_of_missing_policy_returns_not_found()
  {
    using var harness = new Harness { Deleted = true };
    var result = await harness.Service.BeginDeleteAsync(harness.Id, TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.NotFound, result.Error.Type);
    Assert.Contains("no longer exists", result.Error.Message);
  }

  [Theory]
  [InlineData(403, GraphOperationErrorType.Forbidden)]
  [InlineData(429, GraphOperationErrorType.Throttled)]
  [InlineData(503, GraphOperationErrorType.GraphFailure)]
  public async Task DELETE_errors_preserve_diagnostics_and_use_safe_messages(int status, GraphOperationErrorType type)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.DeleteStatus = (HttpStatusCode)status;
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(type, result.Error.Type);
    Assert.Equal(status, result.Error.HttpStatus);
    Assert.Equal("test-error", result.Error.Code);
    Assert.Equal("delete-request", result.Error.RequestId);
    Assert.DoesNotContain("sensitive", result.Error.Message);
    Assert.False(harness.Deleted);
    Assert.Single(harness.Requests, request => request.Method == HttpMethod.Delete);
    if (status == 403)
    {
      Assert.Contains("Policy.ReadWrite.ApplicationConfiguration", result.Error.Message);
      Assert.DoesNotContain("Application.ReadWrite.All", result.Error.Message);
    }
  }

  [Fact]
  public async Task Failed_assignment_read_prevents_DELETE()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.AssignmentStatus = HttpStatusCode.Forbidden;
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.Forbidden, result.Error.Type);
    Assert.DoesNotContain(harness.Requests, request => request.Method == HttpMethod.Delete);
  }

  [Fact]
  public async Task Consent_failure_mentions_existing_permissions_without_exposing_authentication_details()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.BeforeClient = () => throw new MsalUiRequiredException("consent_required", "sensitive authentication");
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.ConsentRequired, result.Error.Type);
    Assert.Contains("Policy.ReadWrite.ApplicationConfiguration", result.Error.Message);
    Assert.DoesNotContain("sensitive", result.Error.Message);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Tenant_switch_invalidates_confirmation_even_after_switching_back(bool switchBack)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    if (switchBack) harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantA);
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Empty(harness.Requests);
  }

  [Theory]
  [InlineData("client")]
  [InlineData("policy")]
  [InlineData("assignments")]
  public async Task Tenant_switch_during_safety_checks_never_issues_DELETE(string stage)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    if (stage == "client") harness.BeforeClient = () => harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    else harness.BeforeSend = (request, _) =>
    {
      if (stage == "policy" || request.RequestUri!.AbsolutePath.EndsWith("/appliesTo"))
        harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      return Task.CompletedTask;
    };
    harness.Requests.Clear();
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.DoesNotContain(harness.Requests, request => request.Method == HttpMethod.Delete);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Network_failure_or_timeout_is_not_reported_as_success(bool timeout)
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    harness.BeforeSend = (request, _) => request.Method == HttpMethod.Delete
      ? throw (timeout ? new OperationCanceledException() : new HttpRequestException())
      : Task.CompletedTask;
    var result = await harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    Assert.Contains("may have completed", result.Error.Message);
    Assert.False(harness.Deleted);
    Assert.Single(harness.Requests, request => request.Method == HttpMethod.Delete);
  }

  [Fact]
  public async Task Caller_cancellation_is_propagated_without_DELETE()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    using var cancellation = new CancellationTokenSource();
    harness.BeforeSend = async (_, token) =>
    {
      cancellation.Cancel();
      await Task.Delay(Timeout.Infinite, token);
    };
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => harness.Service.DeleteAsync(context, "Policy", cancellation.Token));
    Assert.DoesNotContain(harness.Requests, request => request.Method == HttpMethod.Delete);
  }

  [Fact]
  public async Task Late_DELETE_response_never_retrieves_old_ID_in_new_tenant()
  {
    using var harness = new Harness();
    var context = await harness.BeginAsync();
    var pending = new TaskCompletionSource();
    harness.BeforeSend = (request, _) => request.Method == HttpMethod.Delete ? pending.Task : Task.CompletedTask;
    harness.Requests.Clear();
    var deletion = harness.Service.DeleteAsync(context, "Policy", TestContext.Current.CancellationToken);
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    pending.SetResult();
    var result = await deletion;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Equal(HttpMethod.Delete, harness.Requests[^1].Method);
    Assert.Equal(3, harness.Requests.Count);
  }

  private sealed class Harness : IEntraGraphClientFactory, IDisposable
  {
    private readonly DirectoryReadOperation _operation;
    public GraphTestContext Tenants { get; } = new();
    public Guid Id { get; } = Guid.NewGuid();
    public string Path => $"/v1.0/policies/claimsMappingPolicies/{Id:D}";
    public string? Name { get; set; } = "Policy";
    public string Definition { get; set; } = ClaimsMappingPolicyEditingTests.Supported;
    public int AssignmentCount { get; set; }
    public string ObjectType { get; set; } = "servicePrincipal";
    public bool AssignmentOnNextPage { get; set; }
    public bool Deleted { get; set; }
    public HttpStatusCode DeleteStatus { get; set; } = HttpStatusCode.NoContent;
    public HttpStatusCode AssignmentStatus { get; set; } = HttpStatusCode.OK;
    public Action? BeforeClient { get; set; }
    public Func<HttpRequestMessage, CancellationToken, Task>? BeforeSend { get; set; }
    public List<(HttpMethod Method, string Path, string Query, string Body)> Requests { get; } = [];
    public ClaimsMappingPolicyService Service { get; }
    public int ClientCount { get; private set; }

    public Harness()
    {
      Tenants.Tenants.SelectTenant(Tenants.TenantA);
      _operation = new(Tenants.Tenants, this, NullLogger<DirectoryReadOperation>.Instance);
      Service = new(_operation, Tenants.Tenants, this, NullLogger<ClaimsMappingPolicyService>.Instance);
    }

    public async Task<ClaimsMappingPolicyDeletionContext> BeginAsync()
    {
      var result = await Service.BeginDeleteAsync(Id, TestContext.Current.CancellationToken);
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
      var path = request.RequestUri!.AbsolutePath;
      var query = request.RequestUri.Query;
      Requests.Add((request.Method, path, query, request.Content is null ? "" : await request.Content.ReadAsStringAsync(token)));
      Assert.StartsWith(Path, path);
      if (BeforeSend is { } before) await before(request, token);
      if (Deleted) return Error(HttpStatusCode.NotFound);
      if (request.Method == HttpMethod.Delete)
      {
        Assert.Equal(Path, path);
        if (DeleteStatus != HttpStatusCode.NoContent) return Error(DeleteStatus);
        Deleted = true;
        return new(HttpStatusCode.NoContent);
      }
      Assert.Equal(HttpMethod.Get, request.Method);
      if (path.EndsWith("/appliesTo"))
      {
        if (AssignmentStatus != HttpStatusCode.OK) return Error(AssignmentStatus);
        if (AssignmentOnNextPage && !query.Contains("next=1"))
          return Json(new Dictionary<string, object> { ["value"] = Array.Empty<object>(), ["@odata.nextLink"] = $"https://graph.microsoft.com{Path}/appliesTo?next=1" });
        return Json(new
        {
          value = Enumerable.Range(0, AssignmentCount).Select(index => new Dictionary<string, object>
          {
            ["@odata.type"] = $"#microsoft.graph.{ObjectType}",
            ["id"] = Guid.NewGuid(),
            ["appId"] = Guid.NewGuid(),
            ["displayName"] = $"Enterprise Application {index}"
          }).ToArray()
        });
      }
      return Json(new { id = Id, displayName = Name, definition = new[] { Definition }, isOrganizationDefault = false });
    }

    private static HttpResponseMessage Error(HttpStatusCode status) => Json(new
    {
      error = new { code = "test-error", message = "sensitive Graph details", innerError = new Dictionary<string, string> { ["request-id"] = "delete-request" } }
    }, status);
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
