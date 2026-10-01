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

public sealed class ClaimsMappingPolicyAssignmentServiceTests
{
  private static readonly Guid PrincipalId = Guid.Parse("11111111-1111-1111-1111-111111111111");
  private static readonly Guid PolicyId = Guid.Parse("22222222-2222-2222-2222-222222222222");
  private static readonly Guid ClientId = Guid.Parse("33333333-3333-3333-3333-333333333333");
  private static readonly Guid ApplicationId = Guid.Parse("44444444-4444-4444-4444-444444444444");
  private static string PrincipalPath => $"/v1.0/servicePrincipals/{PrincipalId:D}";
  private static string PolicyPath => $"/v1.0/policies/claimsMappingPolicies/{PolicyId:D}";
  private static string AssignedPath => $"{PrincipalPath}/claimsMappingPolicies";

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Mutations_use_principal_object_ID_and_only_the_ref_relationship(bool unassign)
  {
    string? body = null;
    using var harness = new Harness(async (request, token) =>
    {
      if (request.Method == HttpMethod.Get)
        return ExistingObjects(request, unassign);
      Assert.Equal(unassign ? HttpMethod.Delete : HttpMethod.Post, request.Method);
      Assert.Equal(unassign ? $"{AssignedPath}/{PolicyId:D}/$ref" : $"{AssignedPath}/$ref", request.RequestUri!.AbsolutePath);
      Assert.DoesNotContain(ClientId.ToString(), request.RequestUri.ToString());
      Assert.DoesNotContain(ApplicationId.ToString(), request.RequestUri.ToString());
      if (!unassign)
        body = await request.Content!.ReadAsStringAsync(token);
      return new(HttpStatusCode.NoContent);
    });
    var result = unassign
      ? await harness.Service.UnassignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken)
      : await harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(4, harness.Requests.Count);
    if (!unassign)
    {
      using var json = JsonDocument.Parse(body!);
      Assert.Equal($"https://graph.microsoft.com{PolicyPath}", json.RootElement.GetProperty("@odata.id").GetString());
    }
    else
      Assert.Null(harness.Requests[^1].Content);
  }

  [Fact]
  public async Task Direct_discovery_pages_the_principal_navigation_and_preserves_multiple_assignments()
  {
    var second = Guid.NewGuid();
    using var harness = new Harness(request =>
    {
      Assert.Equal(HttpMethod.Get, request.Method);
      Assert.Equal(AssignedPath, request.RequestUri!.AbsolutePath);
      return request.RequestUri.Query.Contains("skiptoken")
        ? Page([new { id = second, displayName = "Second policy" }])
        : Page([new { id = PolicyId, displayName = "First policy" }], $"https://graph.microsoft.com{AssignedPath}?$skiptoken=next");
    });
    var result = await harness.Service.GetAssignedPoliciesAsync(harness.Context, PrincipalId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(new[] { PolicyId, second }, result.Value.Select(item => item.ObjectId));
    Assert.Equal(2, harness.Requests.Count);
    Assert.Contains("$select=id,displayName,definition,isOrganizationDefault", Uri.UnescapeDataString(harness.Requests[0].RequestUri!.Query));
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Stale_duplicate_or_absent_relationship_returns_a_useful_error_without_mutation(bool unassign)
  {
    using var harness = new Harness(request => ExistingObjects(request, assigned: !unassign));
    var result = unassign
      ? await harness.Service.UnassignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken)
      : await harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.InvalidInput, result.Error.Type);
    Assert.Contains(unassign ? "already unassigned" : "already assigned", result.Error.Message);
    Assert.All(harness.Requests, request => Assert.Equal(HttpMethod.Get, request.Method));
  }

  [Theory]
  [InlineData(400, GraphOperationErrorType.InvalidInput)]
  [InlineData(401, GraphOperationErrorType.Unauthorized)]
  [InlineData(403, GraphOperationErrorType.Forbidden)]
  [InlineData(404, GraphOperationErrorType.NotFound)]
  [InlineData(409, GraphOperationErrorType.InvalidInput)]
  [InlineData(429, GraphOperationErrorType.Throttled)]
  [InlineData(503, GraphOperationErrorType.GraphFailure)]
  public async Task Graph_rejections_during_mutation_are_safe_structured_failures(int status, GraphOperationErrorType type)
  {
    using var harness = new Harness(request => request.Method == HttpMethod.Get ? ExistingObjects(request, false)
      : Json(new { error = new { code = "test", message = "sensitive", innerError = new Dictionary<string, string> { ["request-id"] = "assignment-request" } } }, (HttpStatusCode)status));
    var result = await harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(type, result.Error.Type);
    Assert.Equal(status, result.Error.HttpStatus);
    Assert.Equal("assignment-request", result.Error.RequestId);
    Assert.DoesNotContain("sensitive", result.Error.Message);
    if (status == 403)
    {
      Assert.Contains("Application.ReadWrite.All", result.Error.Message);
      Assert.Contains("admin consent", result.Error.Message);
    }
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Old_confirmation_is_rejected_even_after_switching_back(bool unassign)
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not request Graph"));
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantA);
    var result = unassign
      ? await harness.Service.UnassignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken)
      : await harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Equal(0, harness.ClientCount);
    Assert.Empty(harness.Requests);
  }

  [Fact]
  public async Task Tenant_switch_during_client_acquisition_prevents_using_old_object_IDs()
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not request Graph"));
    var pending = new TaskCompletionSource();
    harness.WaitForClient = () => pending.Task;
    var task = harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    pending.SetResult();
    var result = await task;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Empty(harness.Requests);
  }

  [Theory]
  [InlineData(PrincipalPathSuffix.Principal)]
  [InlineData(PrincipalPathSuffix.Policy)]
  [InlineData(PrincipalPathSuffix.Assignments)]
  [InlineData(PrincipalPathSuffix.Mutation)]
  public async Task Tenant_switch_during_each_stage_stops_further_requests_and_discards_late_success(PrincipalPathSuffix stage)
  {
    var pending = new TaskCompletionSource<HttpResponseMessage>();
    using var harness = new Harness((request, _) =>
      (stage == PrincipalPathSuffix.Mutation ? request.Method != HttpMethod.Get : request.RequestUri!.AbsolutePath == (stage switch
      {
        PrincipalPathSuffix.Principal => PrincipalPath,
        PrincipalPathSuffix.Policy => PolicyPath,
        _ => AssignedPath
      })) ? pending.Task : Task.FromResult(ExistingObjects(request, false)));
    var task = harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
    var count = harness.Requests.Count;
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
    harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantA);
    pending.SetResult(stage switch
    {
      PrincipalPathSuffix.Principal => Json(new { id = PrincipalId }),
      PrincipalPathSuffix.Policy => Json(new { id = PolicyId }),
      PrincipalPathSuffix.Assignments => Page([]),
      _ => new(HttpStatusCode.NoContent)
    });
    var result = await task;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Equal(count, harness.Requests.Count);
  }

  public enum PrincipalPathSuffix { Principal, Policy, Assignments, Mutation }

  [Fact]
  public async Task Consent_network_and_timeout_failures_do_not_expose_infrastructure_details()
  {
    foreach (var exception in new Exception[]
    {
      new MsalUiRequiredException("consent_required", "sensitive"),
      new HttpRequestException("sensitive"), new TaskCanceledException("sensitive")
    })
    {
      using var harness = new Harness(_ => throw exception);
      var result = await harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
      Assert.True(result.IsFailure);
      Assert.DoesNotContain("sensitive", result.Error.Message);
      if (exception is MsalUiRequiredException)
      {
        Assert.Equal(GraphOperationErrorType.ConsentRequired, result.Error.Type);
        Assert.Contains("Application.ReadWrite.All", result.Error.Message);
      }
      else
        Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    }
  }

  [Fact]
  public async Task Missing_policy_or_principal_never_mutates_and_invalid_IDs_never_acquire_a_client()
  {
    using var harness = new Harness(_ => Json(new { error = new { code = "deleted", message = "gone" } }, HttpStatusCode.NotFound));
    var invalid = await harness.Service.AssignAsync(harness.Context, Guid.Empty, PolicyId, TestContext.Current.CancellationToken);
    Assert.True(invalid.IsFailure);
    Assert.Equal(GraphOperationErrorType.InvalidInput, invalid.Error.Type);
    Assert.Equal(0, harness.ClientCount);
    var missing = await harness.Service.AssignAsync(harness.Context, PrincipalId, PolicyId, TestContext.Current.CancellationToken);
    Assert.True(missing.IsFailure);
    Assert.Equal(GraphOperationErrorType.NotFound, missing.Error.Type);
    Assert.Single(harness.Requests);
    Assert.Equal(HttpMethod.Get, harness.Requests[0].Method);
  }

  [Fact]
  public async Task Invalid_assignment_next_links_are_rejected_and_cancellation_reaches_the_SDK()
  {
    using var harness = new Harness(_ => Page([], "https://example.test/v1.0/servicePrincipals"));
    var result = await harness.Service.GetAssignedPoliciesAsync(harness.Context, PrincipalId, TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    Assert.Single(harness.Requests);
    using var cancellation = new CancellationTokenSource();
    using var canceled = new Harness(async (_, token) =>
    {
      cancellation.Cancel();
      await Task.Delay(Timeout.Infinite, token);
      throw new InvalidOperationException("Unreachable");
    });
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => canceled.Service.AssignAsync(canceled.Context, PrincipalId, PolicyId, cancellation.Token));
  }

  private static HttpResponseMessage ExistingObjects(HttpRequestMessage request, bool assigned) => request.RequestUri!.AbsolutePath switch
  {
    var path when path == PrincipalPath => Json(new { id = PrincipalId, appId = ClientId }),
    var path when path == PolicyPath => Json(new { id = PolicyId }),
    var path when path == AssignedPath => Page(assigned ? [new { id = PolicyId, displayName = "Policy" }] : []),
    _ => throw new InvalidOperationException("Unexpected Graph request")
  };
  private static HttpResponseMessage Page(object[] items, string? next = null) => Json(new Dictionary<string, object?>
  {
    ["value"] = items, ["@odata.nextLink"] = next
  });
  private static HttpResponseMessage Json(object value, HttpStatusCode status = HttpStatusCode.OK) => new(status)
  {
    Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json")
  };

  private sealed class Harness : IEntraGraphClientFactory, IDisposable
  {
    private readonly DirectoryReadOperation _operation;
    private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _send;
    public GraphTestContext Tenants { get; } = new();
    public ClaimsMappingPolicyService Service { get; }
    public ClaimsMappingPolicyAssignmentContext Context { get; }
    public List<HttpRequestMessage> Requests { get; } = [];
    public int ClientCount { get; private set; }
    public Func<Task>? WaitForClient { get; set; }

    public Harness(Func<HttpRequestMessage, HttpResponseMessage> send) : this((request, _) => Task.FromResult(send(request))) { }
    public Harness(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send)
    {
      _send = send;
      Tenants.Tenants.SelectTenant(Tenants.TenantA);
      _operation = new(Tenants.Tenants, this, NullLogger<DirectoryReadOperation>.Instance);
      Service = new(_operation, Tenants.Tenants, this, NullLogger<ClaimsMappingPolicyService>.Instance);
      Context = Service.BeginAssignment().Value;
    }
    public async Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(CancellationToken cancellationToken = default)
    {
      ClientCount++;
      if (WaitForClient is { } wait) await wait();
      return Result.Success<GraphServiceClient, GraphOperationError>(new GraphServiceClient(
        new HttpClient(new Transport((request, token) =>
        {
          Requests.Add(request);
          return _send(request, token);
        })), new AnonymousAuthenticationProvider()));
    }
    public void Dispose() => _operation.Dispose();
  }
  private sealed class Transport(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) : HttpMessageHandler
  {
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => send(request, cancellationToken);
  }
}
