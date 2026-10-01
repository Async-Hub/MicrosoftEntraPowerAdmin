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

public sealed class ClaimsMappingPolicyServiceTests
{
  private static readonly Guid PolicyId = Guid.Parse("11111111-1111-1111-1111-111111111111");
  private static readonly Guid PrincipalId = Guid.Parse("22222222-2222-2222-2222-222222222222");
  private static readonly Guid AppId = Guid.Parse("33333333-3333-3333-3333-333333333333");
  private static readonly Guid ApplicationId = Guid.Parse("44444444-4444-4444-4444-444444444444");
  private const string Collection = "/v1.0/policies/claimsMappingPolicies";

  [Fact]
  public async Task Lists_all_pages_with_parsed_summaries_and_isolates_bad_definitions()
  {
    using var harness = new Harness(request => request.RequestUri?.Query.Contains("skiptoken") == true
      ? Page([Policy(Guid.NewGuid(), "Malformed", ["broken"]), Policy(Guid.NewGuid(), "Empty", [])])
      : Page([Policy(PolicyId, "Department", [ClaimsMappingDefinitionParserTests.Definition])],
        $"https://graph.microsoft.com{Collection}?$skiptoken=next"));
    var result = await harness.Service.ListAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(3, result.Value.Count);
    var first = result.Value[0];
    Assert.Equal(PolicyId, first.ObjectId);
    Assert.Equal(2, first.ClaimsSchemaCount);
    Assert.Equal(1, first.ClaimsTransformationCount);
    Assert.False(first.IsOrganizationDefault);
    Assert.NotEmpty(result.Value[1].Definitions[0].Warnings);
    Assert.False(result.Value[2].HasDefinition);
    Assert.Equal(2, harness.Requests.Count);
    Assert.Contains("$select=id,displayName,definition,isOrganizationDefault", Query(harness.Requests[0]));
  }

  [Fact]
  public async Task Details_page_all_assignments_and_keep_the_three_ID_types_distinct()
  {
    using var harness = new Harness(request =>
    {
      var path = request.RequestUri!.AbsolutePath;
      if (path == $"{Collection}/{PolicyId}")
        return Json(Policy(PolicyId, "Department", [ClaimsMappingDefinitionParserTests.Definition]));
      Assert.Equal($"{Collection}/{PolicyId}/appliesTo", path);
      if (request.RequestUri.Query.Contains("skiptoken"))
        return Page([DirectoryObject("application", ApplicationId, AppId, "Registration")]);
      return Page([DirectoryObject("servicePrincipal", PrincipalId, AppId, "Enterprise")],
        $"https://graph.microsoft.com{path}?$skiptoken=next");
    });
    var result = await harness.Service.GetByObjectIdAsync(PolicyId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.True(result.Value.Assignments.IsSuccess);
    var assignments = result.Value.Assignments.Value;
    var principal = Assert.Single(assignments.ServicePrincipals);
    Assert.Equal(PrincipalId, principal.ObjectId);
    Assert.Equal(AppId, principal.AppId);
    var application = Assert.Single(assignments.Applications);
    Assert.Equal(ApplicationId, application.ObjectId);
    Assert.Equal(AppId, application.AppId);
    Assert.Equal(3, harness.Requests.Count);
  }

  [Fact]
  public async Task Reverse_discovery_matches_only_service_principal_object_ID()
  {
    var matchingPolicy = Guid.NewGuid();
    var clientIdMatchPolicy = Guid.NewGuid();
    using var harness = new Harness(request =>
    {
      var path = request.RequestUri!.AbsolutePath;
      if (path == Collection)
        return Page([Policy(PolicyId, "Application object match", []), Policy(clientIdMatchPolicy, "Client ID match", []), Policy(matchingPolicy, "Actual match", [])]);
      if (path == $"{Collection}/{PolicyId}/appliesTo")
        return Page([DirectoryObject("application", PrincipalId, AppId, "Same application object ID")]);
      if (path == $"{Collection}/{clientIdMatchPolicy}/appliesTo")
        return Page([DirectoryObject("servicePrincipal", ApplicationId, PrincipalId, "Same Client ID")]);
      Assert.Equal($"{Collection}/{matchingPolicy}/appliesTo", path);
      return Page([DirectoryObject("servicePrincipal", PrincipalId, AppId, "Actual assignment")]);
    });
    var result = await harness.Service.FindForServicePrincipalAsync(PrincipalId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(matchingPolicy, Assert.Single(result.Value).ObjectId);
    Assert.Equal(4, harness.Requests.Count);
  }

  [Fact]
  public async Task Policy_and_relationship_pages_do_not_hide_assignments()
  {
    var secondPolicy = Guid.NewGuid();
    using var harness = new Harness(request =>
    {
      var path = request.RequestUri!.AbsolutePath;
      if (path == Collection && request.RequestUri.Query.Contains("skiptoken"))
        return Page([Policy(secondPolicy, "Second policy", [])]);
      if (path == Collection)
        return Page([Policy(PolicyId, "First policy", [])], $"https://graph.microsoft.com{Collection}?$skiptoken=policies");
      if (path == $"{Collection}/{secondPolicy}/appliesTo")
        return Page([]);
      Assert.Equal($"{Collection}/{PolicyId}/appliesTo", path);
      return request.RequestUri.Query.Contains("skiptoken")
        ? Page([DirectoryObject("servicePrincipal", PrincipalId, AppId, "Assigned on next page")])
        : Page([], $"https://graph.microsoft.com{path}?$skiptoken=assignments");
    });
    var result = await harness.Service.FindForServicePrincipalAsync(PrincipalId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(PolicyId, Assert.Single(result.Value).ObjectId);
    Assert.Equal(5, harness.Requests.Count);
  }

  [Theory]
  [InlineData(401, GraphOperationErrorType.Unauthorized)]
  [InlineData(403, GraphOperationErrorType.Forbidden)]
  [InlineData(404, GraphOperationErrorType.NotFound)]
  [InlineData(429, GraphOperationErrorType.Throttled)]
  [InlineData(503, GraphOperationErrorType.GraphFailure)]
  public async Task Expected_Graph_errors_have_safe_messages_and_preserve_metadata(int status, GraphOperationErrorType type)
  {
    using var harness = new Harness(_ => Json(new
    {
      error = new { code = "test", message = "sensitive Graph error", innerError = new Dictionary<string, string> { ["request-id"] = "request-1" } }
    }, (HttpStatusCode)status));
    var result = await harness.Service.ListAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(type, result.Error.Type);
    Assert.Equal(status, result.Error.HttpStatus);
    Assert.Equal("request-1", result.Error.RequestId);
    Assert.DoesNotContain("sensitive", result.Error.Message);
    if (status == 403)
    {
      Assert.Contains("Policy.Read.All", result.Error.Message);
      Assert.Contains("Application.Read.All", result.Error.Message);
      Assert.DoesNotContain("ReadWrite", result.Error.Message);
    }
  }

  [Fact]
  public async Task Missing_consent_identifies_read_only_permissions()
  {
    using var harness = new Harness(_ => throw new MsalUiRequiredException("consent_required", "sensitive message"));
    var result = await harness.Service.ListAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.ConsentRequired, result.Error.Type);
    Assert.Contains("Policy.Read.All", result.Error.Message);
    Assert.DoesNotContain("sensitive", result.Error.Message);
  }

  [Fact]
  public async Task Assignment_failure_preserves_the_policy_definition()
  {
    using var harness = new Harness(request => request.RequestUri!.AbsolutePath.EndsWith("appliesTo")
      ? Json(new { error = new { code = "deleted", message = "removed" } }, HttpStatusCode.NotFound)
      : Json(Policy(PolicyId, "Policy", [ClaimsMappingDefinitionParserTests.Definition])));
    var result = await harness.Service.GetByObjectIdAsync(PolicyId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(ClaimsMappingDefinitionParserTests.Definition, Assert.Single(result.Value.Policy.Definitions).Raw);
    Assert.True(result.Value.Assignments.IsFailure);
    Assert.Equal(GraphOperationErrorType.NotFound, result.Value.Assignments.Error.Type);
  }

  [Fact]
  public async Task Missing_definition_and_empty_assignments_are_normal_results()
  {
    using var harness = new Harness(request => request.RequestUri!.AbsolutePath.EndsWith("appliesTo")
      ? Page([]) : Json(new { id = PolicyId, displayName = "No definition" }));
    var result = await harness.Service.GetByObjectIdAsync(PolicyId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.False(result.Value.Policy.HasDefinition);
    Assert.Empty(result.Value.Policy.Definitions);
    Assert.True(result.Value.Assignments.IsSuccess);
    Assert.Empty(result.Value.Assignments.Value.ServicePrincipals);
  }

  [Theory]
  [InlineData("https://example.test/v1.0/policies/claimsMappingPolicies")]
  [InlineData("https://graph.microsoft.com/beta/policies/claimsMappingPolicies")]
  [InlineData("https://graph.microsoft.com/v1.0/servicePrincipals")]
  public async Task Invalid_next_links_are_rejected_before_another_request(string next)
  {
    using var harness = new Harness(_ => Page([], next));
    var result = await harness.Service.ListAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    Assert.Single(harness.Requests);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Tenant_switch_stops_paging_and_assignment_requests_even_after_switching_back(bool details)
  {
    var pending = new TaskCompletionSource<HttpResponseMessage>();
    using var harness = new Harness((_, _) => pending.Task);
    var task = details
      ? (Task)harness.Service.GetByObjectIdAsync(PolicyId, TestContext.Current.CancellationToken)
      : harness.Service.FindForServicePrincipalAsync(PrincipalId, TestContext.Current.CancellationToken);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantA);
    pending.SetResult(details ? Json(Policy(PolicyId, "Stale", [])) :
      Page([Policy(PolicyId, "Stale", [])], $"https://graph.microsoft.com{Collection}?$skiptoken=next"));
    var error = details
      ? (await (Task<Result<ClaimsMappingPolicyDetails, GraphOperationError>>)task).Error
      : (await (Task<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>>)task).Error;
    Assert.Equal(GraphOperationErrorType.TenantChanged, error.Type);
    Assert.Single(harness.Requests);
  }

  [Fact]
  public async Task Cancellation_is_propagated_to_Graph()
  {
    using var cancellation = new CancellationTokenSource();
    using var harness = new Harness(async (_, token) =>
    {
      cancellation.Cancel();
      await Task.Delay(Timeout.Infinite, token);
      throw new InvalidOperationException("Unreachable");
    });
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => harness.Service.ListAsync(cancellation.Token));
  }

  [Fact]
  public async Task Tenant_switch_during_assignment_read_prevents_reading_more_old_policy_IDs()
  {
    var pending = new TaskCompletionSource<HttpResponseMessage>();
    using var harness = new Harness((request, _) => request.RequestUri!.AbsolutePath == Collection
      ? Task.FromResult(Page([Policy(PolicyId, "First policy", []), Policy(Guid.NewGuid(), "Second policy", [])])) : pending.Task);
    var task = harness.Service.FindForServicePrincipalAsync(PrincipalId, TestContext.Current.CancellationToken);
    Assert.Equal(2, harness.Requests.Count);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    pending.SetResult(Page([DirectoryObject("servicePrincipal", PrincipalId, AppId, "Stale assignment")]));
    var result = await task;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Equal(2, harness.Requests.Count);
  }

  [Fact]
  public async Task No_selected_tenant_prevents_client_creation()
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not query Graph"), selectTenant: false);
    var result = await harness.Service.ListAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantNotSelected, result.Error.Type);
    Assert.Equal(0, harness.ClientCount);
  }

  [Fact]
  public async Task Empty_IDs_do_not_make_Graph_requests()
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not query Graph"));
    var invalid = await harness.Service.GetByObjectIdAsync(Guid.Empty, TestContext.Current.CancellationToken);
    Assert.Equal(GraphOperationErrorType.InvalidInput, invalid.Error.Type);
    var reverse = await harness.Service.FindForServicePrincipalAsync(Guid.Empty, TestContext.Current.CancellationToken);
    Assert.Equal(GraphOperationErrorType.InvalidInput, reverse.Error.Type);
    Assert.Empty(harness.Requests);
  }

  [Fact]
  public async Task Creation_posts_one_definition_string_and_uses_Graph_response_without_assigning()
  {
    string? submitted = null;
    var returnedDefinition = "{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"false\",\"ClaimsSchema\":[]}}";
    using var harness = new Harness(async (request, token) =>
    {
      Assert.Equal(HttpMethod.Post, request.Method);
      Assert.Equal(Collection, request.RequestUri!.AbsolutePath);
      submitted = await request.Content!.ReadAsStringAsync(token);
      return Json(Policy(PolicyId, "Name returned by Graph", [returnedDefinition]), HttpStatusCode.Created);
    }, allowCreate: true);
    var context = harness.Service.BeginCreate();
    Assert.True(context.IsSuccess);
    Assert.Equal(harness.Context.TenantA, context.Value.TenantId);
    var result = await harness.Service.CreateAsync(context.Value, CreateRequest(), TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(PolicyId, result.Value.ObjectId);
    Assert.Equal("Name returned by Graph", result.Value.DisplayName);
    Assert.Equal(returnedDefinition, Assert.Single(result.Value.Definitions).Raw);
    Assert.Single(harness.Requests);
    using var outer = JsonDocument.Parse(submitted!);
    Assert.Equal("Policy", outer.RootElement.GetProperty("displayName").GetString());
    var definition = Assert.Single(outer.RootElement.GetProperty("definition").EnumerateArray());
    Assert.Equal(JsonValueKind.String, definition.ValueKind);
    using var inner = JsonDocument.Parse(definition.GetString()!);
    Assert.Equal("department", inner.RootElement.GetProperty("ClaimsMappingPolicy").GetProperty("ClaimsSchema")[0].GetProperty("ID").GetString());
    Assert.False(outer.RootElement.GetProperty("isOrganizationDefault").GetBoolean());
    Assert.False(outer.RootElement.TryGetProperty("appliesTo", out _));
  }

  [Theory]
  [InlineData(400, GraphOperationErrorType.InvalidInput)]
  [InlineData(401, GraphOperationErrorType.Unauthorized)]
  [InlineData(403, GraphOperationErrorType.Forbidden)]
  [InlineData(409, GraphOperationErrorType.InvalidInput)]
  [InlineData(429, GraphOperationErrorType.Throttled)]
  [InlineData(503, GraphOperationErrorType.GraphFailure)]
  public async Task Creation_errors_are_safe_and_preserve_diagnostic_metadata(int status, GraphOperationErrorType type)
  {
    using var harness = new Harness(_ => Json(new
    {
      error = new { code = "test", message = "sensitive message", innerError = new Dictionary<string, string> { ["request-id"] = "create-request" } }
    }, (HttpStatusCode)status), allowCreate: true);
    var result = await harness.Service.CreateAsync(harness.Service.BeginCreate().Value, CreateRequest(), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(type, result.Error.Type);
    Assert.Equal(status, result.Error.HttpStatus);
    Assert.Equal("test", result.Error.Code);
    Assert.Equal("create-request", result.Error.RequestId);
    Assert.DoesNotContain("sensitive", result.Error.Message);
    if (status == 403)
      Assert.Contains("Policy.ReadWrite.ApplicationConfiguration", result.Error.Message);
  }

  [Fact]
  public async Task Creation_consent_failure_identifies_required_write_permission()
  {
    using var harness = new Harness(_ => throw new MsalUiRequiredException("consent_required", "sensitive message"), allowCreate: true);
    var result = await harness.Service.CreateAsync(harness.Service.BeginCreate().Value, CreateRequest(), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.ConsentRequired, result.Error.Type);
    Assert.Contains("Policy.ReadWrite.ApplicationConfiguration", result.Error.Message);
    Assert.DoesNotContain("sensitive", result.Error.Message);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Old_draft_cannot_create_after_tenant_switch_even_if_switched_back(bool switchBack)
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not POST"));
    var context = harness.Service.BeginCreate().Value;
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    if (switchBack) harness.Context.Tenants.SelectTenant(harness.Context.TenantA);
    var result = await harness.Service.CreateAsync(context, CreateRequest(), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Empty(harness.Requests);
    Assert.Equal(0, harness.ClientCount);
  }

  [Fact]
  public async Task Tenant_switch_during_client_acquisition_prevents_POST()
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not POST"));
    var pending = new TaskCompletionSource();
    harness.WaitForClient = () => pending.Task;
    var task = harness.Service.CreateAsync(harness.Service.BeginCreate().Value, CreateRequest(), TestContext.Current.CancellationToken);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    pending.SetResult();
    var result = await task;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Empty(harness.Requests);
  }

  [Fact]
  public async Task Late_creation_response_does_not_leak_into_new_tenant()
  {
    var pending = new TaskCompletionSource<HttpResponseMessage>();
    using var harness = new Harness((_, _) => pending.Task, allowCreate: true);
    var task = harness.Service.CreateAsync(harness.Service.BeginCreate().Value, CreateRequest(), TestContext.Current.CancellationToken);
    Assert.Single(harness.Requests);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantA);
    pending.SetResult(Json(Policy(PolicyId, "Old tenant", []), HttpStatusCode.Created));
    var result = await task;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
    Assert.Single(harness.Requests);
  }

  [Fact]
  public async Task Invalid_draft_is_rejected_before_client_acquisition()
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not POST"));
    var result = await harness.Service.CreateAsync(harness.Service.BeginCreate().Value,
      CreateRequest() with { DisplayName = " " }, TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.InvalidInput, result.Error.Type);
    Assert.Contains("Display Name", result.Error.Message);
    Assert.Equal(0, harness.ClientCount);
  }

  [Fact]
  public void Creation_requires_a_selected_tenant()
  {
    using var harness = new Harness(_ => throw new InvalidOperationException("Must not POST"), selectTenant: false);
    var result = harness.Service.BeginCreate();
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantNotSelected, result.Error.Type);
    Assert.Equal(0, harness.ClientCount);
  }

  [Fact]
  public async Task Missing_creation_response_tells_user_to_check_list_before_retrying()
  {
    using var harness = new Harness(_ => new HttpResponseMessage(HttpStatusCode.NoContent), allowCreate: true);
    var result = await harness.Service.CreateAsync(harness.Service.BeginCreate().Value, CreateRequest(), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    Assert.Contains("Refresh the list", result.Error.Message);
  }

  [Fact]
  public async Task Creation_network_failure_is_a_safe_result()
  {
    using var harness = new Harness(_ => throw new HttpRequestException("sensitive address"), allowCreate: true);
    var result = await harness.Service.CreateAsync(harness.Service.BeginCreate().Value, CreateRequest(), TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    Assert.Contains("refresh the list", result.Error.Message);
    Assert.DoesNotContain("sensitive", result.Error.Message);
  }

  [Fact]
  public async Task Creation_propagates_caller_cancellation_to_Graph()
  {
    using var cancellation = new CancellationTokenSource();
    using var harness = new Harness(async (_, token) =>
    {
      cancellation.Cancel();
      await Task.Delay(Timeout.Infinite, token);
      throw new InvalidOperationException("Unreachable");
    }, allowCreate: true);
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => harness.Service.CreateAsync(
      harness.Service.BeginCreate().Value, CreateRequest(), cancellation.Token));
  }

  private static CreateClaimsMappingPolicyRequest CreateRequest() => new("Policy", true,
    [new(ClaimValueMode.DirectoryAttribute, "user", "department", "department", null, null)]);

  private static Dictionary<string, object> Policy(Guid id, string name, string[] definitions) => new()
  {
    ["id"] = id, ["displayName"] = name, ["definition"] = definitions, ["isOrganizationDefault"] = false
  };
  private static Dictionary<string, object> DirectoryObject(string type, Guid id, Guid appId, string name) => new()
  {
    ["@odata.type"] = $"#microsoft.graph.{type}", ["id"] = id, ["appId"] = appId, ["displayName"] = name
  };
  private static HttpResponseMessage Page(object[] objects, string? next = null) => Json(new Dictionary<string, object?>
  {
    ["value"] = objects, ["@odata.nextLink"] = next
  });
  private static HttpResponseMessage Json(object value, HttpStatusCode status = HttpStatusCode.OK) => new(status)
  {
    Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json")
  };
  private static string Query(HttpRequestMessage request) => Uri.UnescapeDataString(request.RequestUri?.Query ?? "");

  private sealed class Harness : IEntraGraphClientFactory, IDisposable
  {
    private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _send;
    private readonly DirectoryReadOperation _operation;
    public GraphTestContext Context { get; } = new();
    public List<HttpRequestMessage> Requests { get; } = [];
    public ClaimsMappingPolicyService Service { get; }
    public int ClientCount { get; private set; }
    public Func<Task>? WaitForClient { get; set; }
    private readonly bool _allowCreate;

    public Harness(Func<HttpRequestMessage, HttpResponseMessage> send, bool selectTenant = true, bool allowCreate = false)
      : this((request, _) => Task.FromResult(send(request)), selectTenant, allowCreate) { }
    public Harness(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send, bool selectTenant = true, bool allowCreate = false)
    {
      _send = send;
      _allowCreate = allowCreate;
      if (selectTenant) Context.Tenants.SelectTenant(Context.TenantA);
      _operation = new(Context.Tenants, this, NullLogger<DirectoryReadOperation>.Instance);
      Service = new(_operation, Context.Tenants, this, NullLogger<ClaimsMappingPolicyService>.Instance);
    }
    public async Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(CancellationToken cancellationToken = default)
    {
      ClientCount++;
      if (WaitForClient is { } wait) await wait();
      return Result.Success<GraphServiceClient, GraphOperationError>(
        new GraphServiceClient(new HttpClient(new Transport((request, token) =>
        {
          Requests.Add(request);
          Assert.Equal(_allowCreate ? HttpMethod.Post : HttpMethod.Get, request.Method);
          Assert.StartsWith(Collection, request.RequestUri!.AbsolutePath);
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
