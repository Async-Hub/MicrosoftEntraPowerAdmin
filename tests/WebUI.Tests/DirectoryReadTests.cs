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

public sealed class DirectoryReadTests
{
  private static readonly Guid ObjectId = Guid.Parse("11111111-1111-1111-1111-111111111111");
  private static readonly Guid ClientId = Guid.Parse("22222222-2222-2222-2222-222222222222");
  private const string ObjectJson = """
    {"id":"11111111-1111-1111-1111-111111111111","appId":"22222222-2222-2222-2222-222222222222",
     "displayName":"My API","description":"Description","signInAudience":"AzureADMultipleOrgs",
     "createdDateTime":"2026-01-02T03:04:05Z","publisherDomain":"example.test",
     "identifierUris":["api://my-api"],"web":{"redirectUris":["https://example.test/web"]},
     "spa":{"redirectUris":["https://example.test/spa"]},"publicClient":{"redirectUris":["http://localhost"]},
     "api":{"acceptMappedClaims":false,"requestedAccessTokenVersion":2},"tags":["tag"],
     "servicePrincipalType":"Application","accountEnabled":true,
     "verifiedPublisher":{"displayName":"Verified publisher"},"homepage":"https://example.test",
     "loginUrl":"https://example.test/login","servicePrincipalNames":["api://my-api"],
     "appOwnerOrganizationId":"33333333-3333-3333-3333-333333333333"}
    """;

  [Fact]
  public async Task Maps_application_details_and_uses_application_object_id()
  {
    using var harness = new Harness((request, _) =>
    {
      Assert.Equal($"/v1.0/applications/{ObjectId}", request.RequestUri?.AbsolutePath);
      Assert.Contains("api", Query(request));
      return Task.FromResult(Json(ObjectJson));
    });
    var result = await harness.Applications.GetByObjectIdAsync(ObjectId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    var details = result.Value;
    Assert.Equal(ObjectId, details.Application.ObjectId);
    Assert.Equal(ClientId, details.Application.AppId);
    Assert.Equal("My API", details.Application.DisplayName);
    Assert.Equal("AzureADMultipleOrgs", details.Application.SignInAudience);
    Assert.Equal("example.test", details.Application.PublisherDomain);
    Assert.Equal(DateTimeOffset.Parse("2026-01-02T03:04:05Z"), details.Application.CreatedDateTime);
    Assert.Equal("Description", details.Description);
    Assert.Equal(["api://my-api"], details.IdentifierUris);
    Assert.Equal(["https://example.test/web"], details.WebRedirectUris);
    Assert.Equal(["https://example.test/spa"], details.SpaRedirectUris);
    Assert.Equal(["http://localhost"], details.PublicClientRedirectUris);
    Assert.Equal(["tag"], details.Tags);
    Assert.False(details.AcceptMappedClaims);
    Assert.Equal(2, details.RequestedAccessTokenVersion);
  }

  [Theory]
  [InlineData("null", null)]
  [InlineData("false", false)]
  [InlineData("true", true)]
  public async Task Preserves_nullable_API_settings(string jsonValue, bool? expected)
  {
    using var harness = new Harness((_, _) => Task.FromResult(Json(
      JsonSerializer.Serialize(new { id = ObjectId, api = new { acceptMappedClaims = JsonSerializer.Deserialize<bool?>(jsonValue) } }))));
    var result = await harness.Applications.GetByObjectIdAsync(ObjectId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(expected, result.Value.AcceptMappedClaims);
    Assert.Null(result.Value.RequestedAccessTokenVersion);
    Assert.Empty(result.Value.WebRedirectUris);
  }

  [Fact]
  public async Task Maps_service_principal_details_and_uses_service_principal_object_id()
  {
    using var harness = new Harness((request, _) =>
    {
      Assert.Equal($"/v1.0/servicePrincipals/{ObjectId}", request.RequestUri?.AbsolutePath);
      return Task.FromResult(Json(ObjectJson));
    });
    var result = await harness.Principals.GetByObjectIdAsync(ObjectId, TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    var details = result.Value;
    Assert.Equal(ObjectId, details.ServicePrincipal.ObjectId);
    Assert.Equal(ClientId, details.ServicePrincipal.AppId);
    Assert.Equal("My API", details.ServicePrincipal.DisplayName);
    Assert.Equal("Application", details.ServicePrincipal.ServicePrincipalType);
    Assert.True(details.ServicePrincipal.AccountEnabled);
    Assert.Equal("Verified publisher", details.ServicePrincipal.PublisherName);
    Assert.Equal("Description", details.Description);
    Assert.Equal("https://example.test", details.Homepage);
    Assert.Equal("https://example.test/login", details.LoginUrl);
    Assert.Equal(["api://my-api"], details.ServicePrincipalNames);
    Assert.Equal(Guid.Parse("33333333-3333-3333-3333-333333333333"), details.AppOwnerOrganizationId);
    Assert.Equal(["tag"], details.Tags);
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task Searches_display_name_on_server_and_escapes_apostrophes(bool application)
  {
    using var harness = new Harness((request, _) =>
    {
      Assert.Contains("$filter=startswith(displayName,'O''Brien')", Query(request));
      Assert.Contains("$top=25", Query(request));
      Assert.Contains("$select=", Query(request));
      Assert.DoesNotContain("redirectUris", Query(request));
      Assert.Contains("eventual", request.Headers.GetValues("ConsistencyLevel"));
      return Task.FromResult(Json("{\"value\":[" + ObjectJson + "]}"));
    });
    if (application)
    {
      var result = await harness.Applications.SearchAsync(" O'Brien ", cancellationToken: TestContext.Current.CancellationToken);
      Assert.True(result.IsSuccess);
      Assert.Equal(ClientId, Assert.Single(result.Value.Items).AppId);
      Assert.Null(result.Value.NextPage);
    }
    else
    {
      var result = await harness.Principals.SearchAsync(" O'Brien ", cancellationToken: TestContext.Current.CancellationToken);
      Assert.True(result.IsSuccess);
      Assert.Equal(ClientId, Assert.Single(result.Value.Items).AppId);
      Assert.Null(result.Value.NextPage);
    }
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task GUID_search_matches_client_id_or_object_id(bool application)
  {
    using var harness = new Harness((request, _) =>
    {
      Assert.Contains($"$filter=appId eq '{ClientId}' or id eq '{ClientId}'", Query(request));
      return Task.FromResult(Json("{\"value\":[]}"));
    });
    Assert.True((await harness.SearchAsync(application, ClientId.ToString())).IsSuccess);
  }

  [Theory]
  [InlineData(true, true)]
  [InlineData(true, false)]
  [InlineData(false, true)]
  [InlineData(false, false)]
  public async Task Exact_client_id_lookup_returns_match_or_expected_absence(bool application, bool found)
  {
    using var harness = new Harness((request, _) =>
    {
      Assert.Contains($"$filter=appId eq '{ClientId}'", Query(request));
      Assert.DoesNotContain(ObjectId.ToString(), Query(request));
      Assert.Contains("$top=2", Query(request));
      return Task.FromResult(Json("{\"value\":[" + (found ? ObjectJson : "") + "]}"));
    });
    if (application)
    {
      var result = await harness.Applications.FindByAppIdAsync(ClientId, TestContext.Current.CancellationToken);
      Assert.True(result.IsSuccess);
      Assert.Equal(found, result.Value.HasValue);
      if (found) Assert.Equal(ObjectId, result.Value.Value.Application.ObjectId);
    }
    else
    {
      var result = await harness.Principals.FindByAppIdAsync(ClientId, TestContext.Current.CancellationToken);
      Assert.True(result.IsSuccess);
      Assert.Equal(found, result.Value.HasValue);
      if (found) Assert.Equal(ObjectId, result.Value.Value.ServicePrincipal.ObjectId);
    }
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task Paging_follows_SDK_next_link_and_rejects_it_after_tenant_switch(bool application)
  {
    var collection = application ? "applications" : "servicePrincipals";
    var nextUrl = $"https://graph.microsoft.com/v1.0/{collection}?$skiptoken=opaque%2Btoken";
    var calls = 0;
    using var harness = new Harness((request, _) =>
    {
      calls++;
      if (calls == 1)
        return Task.FromResult(Json(JsonSerializer.Serialize(new Dictionary<string, object>
        {
          ["value"] = Array.Empty<object>(),
          ["@odata.nextLink"] = nextUrl
        })));
      Assert.Equal(nextUrl, request.RequestUri?.AbsoluteUri);
      Assert.Contains("eventual", request.Headers.GetValues("ConsistencyLevel"));
      return Task.FromResult(Json("{\"value\":[]}"));
    });
    var first = await harness.SearchAsync(application);
    Assert.True(first.IsSuccess);
    Assert.NotNull(first.Value);
    var second = await harness.SearchAsync(application, next: first.Value);
    Assert.True(second.IsSuccess);
    Assert.Null(second.Value);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantA);
    var stale = await harness.SearchAsync(application, next: first.Value);
    Assert.True(stale.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, stale.Error.Type);
    Assert.Equal(2, calls);
  }

  [Theory]
  [InlineData("https://example.test/v1.0/applications")]
  [InlineData("http://graph.microsoft.com/v1.0/applications")]
  [InlineData("https://graph.microsoft.com/beta/applications")]
  [InlineData("https://graph.microsoft.com/v1.0/users")]
  public async Task Rejects_next_links_outside_the_expected_collection(string url)
  {
    using var harness = new Harness((_, _) => Task.FromResult(Json(JsonSerializer.Serialize(new Dictionary<string, object>
    {
      ["value"] = Array.Empty<object>(),
      ["@odata.nextLink"] = url
    }))));
    var result = await harness.SearchAsync(true);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task Every_search_resolves_current_tenant_and_does_not_reuse_results(bool application)
  {
    var name = "Tenant A application";
    using var harness = new Harness((_, _) => Task.FromResult(Json(
      "{\"value\":[" + ObjectJson.Replace("My API", name) + "]}")));
    Assert.Equal("Tenant A application", await SearchNameAsync());
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    name = "Tenant B application";
    Assert.Equal("Tenant B application", await SearchNameAsync());
    Assert.Equal([harness.Context.TenantA, harness.Context.TenantB], harness.CreatedForTenants);

    async Task<string> SearchNameAsync()
    {
      if (application)
      {
        var result = await harness.Applications.SearchAsync(null, cancellationToken: TestContext.Current.CancellationToken);
        Assert.True(result.IsSuccess);
        return Assert.Single(result.Value.Items).DisplayName;
      }
      var principals = await harness.Principals.SearchAsync(null, cancellationToken: TestContext.Current.CancellationToken);
      Assert.True(principals.IsSuccess);
      return Assert.Single(principals.Value.Items).DisplayName;
    }
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task Late_search_response_is_rejected_even_after_switching_back_to_A(bool application)
  {
    var pending = new TaskCompletionSource<HttpResponseMessage>();
    using var harness = new Harness((_, _) => pending.Task);
    var task = harness.SearchAsync(application);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
    harness.Context.Tenants.SelectTenant(harness.Context.TenantA);
    pending.SetResult(Json("{\"value\":[" + ObjectJson + "]}"));
    var result = await task;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task Cancellation_reaches_Graph_transport(bool application)
  {
    using var cancellation = new CancellationTokenSource();
    using var harness = new Harness(async (_, token) =>
    {
      Assert.True(token.CanBeCanceled);
      cancellation.Cancel();
      await Task.Delay(Timeout.Infinite, token);
      throw new InvalidOperationException("Unreachable");
    });
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => application
      ? (Task)harness.Applications.SearchAsync(null, cancellationToken: cancellation.Token)
      : harness.Principals.SearchAsync(null, cancellationToken: cancellation.Token));
  }

  [Theory]
  [InlineData(401, GraphOperationErrorType.Unauthorized)]
  [InlineData(403, GraphOperationErrorType.Forbidden)]
  [InlineData(404, GraphOperationErrorType.NotFound)]
  [InlineData(429, GraphOperationErrorType.Throttled)]
  [InlineData(500, GraphOperationErrorType.GraphFailure)]
  public async Task Read_errors_are_safe_and_preserve_metadata(int status, GraphOperationErrorType type)
  {
    using var harness = new Harness((_, _) => Task.FromResult(Json("""
      {"error":{"code":"test","message":"sensitive raw exception","innerError":{"request-id":"request-1"}}}
      """, (HttpStatusCode)status)));
    var result = await harness.SearchAsync(true);
    Assert.True(result.IsFailure);
    Assert.Equal(type, result.Error.Type);
    Assert.Equal("request-1", result.Error.RequestId);
    Assert.DoesNotContain("sensitive", result.Error.Message);
    if (status == 403) Assert.Contains("Application.Read.All", result.Error.Message);
  }

  [Theory]
  [InlineData("consent_required", GraphOperationErrorType.ConsentRequired)]
  [InlineData("invalid_grant", GraphOperationErrorType.InteractionRequired)]
  public async Task Authentication_failures_use_existing_error_model(string code, GraphOperationErrorType type)
  {
    using var harness = new Harness((_, _) => throw new MsalUiRequiredException(code, "sensitive"));
    var result = await harness.SearchAsync(false);
    Assert.True(result.IsFailure);
    Assert.Equal(type, result.Error.Type);
  }

  [Fact]
  public async Task No_tenant_selected_does_not_create_client()
  {
    using var harness = new Harness((_, _) => throw new InvalidOperationException(), selectTenant: false);
    var result = await harness.SearchAsync(true);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantNotSelected, result.Error.Type);
    Assert.Empty(harness.CreatedForTenants);
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task Empty_object_id_is_rejected_before_Graph_request(bool application)
  {
    using var harness = new Harness((_, _) => throw new InvalidOperationException("Must not query Graph"));
    var error = application
      ? (await harness.Applications.GetByObjectIdAsync(Guid.Empty, TestContext.Current.CancellationToken)).Error
      : (await harness.Principals.GetByObjectIdAsync(Guid.Empty, TestContext.Current.CancellationToken)).Error;
    Assert.Equal(GraphOperationErrorType.InvalidInput, error.Type);
  }

  private static string Query(HttpRequestMessage request) => Uri.UnescapeDataString(request.RequestUri?.Query ?? "");
  private static HttpResponseMessage Json(string json, HttpStatusCode status = HttpStatusCode.OK) => new(status)
  {
    Content = new StringContent(json, Encoding.UTF8, "application/json")
  };

  private sealed class Harness : IEntraGraphClientFactory, IDisposable
  {
    private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _send;
    private readonly DirectoryReadOperation _operation;
    public GraphTestContext Context { get; } = new();
    public List<Guid> CreatedForTenants { get; } = [];
    public ApplicationService Applications { get; }
    public ServicePrincipalService Principals { get; }

    public Harness(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send, bool selectTenant = true)
    {
      _send = send;
      if (selectTenant) Context.Tenants.SelectTenant(Context.TenantA);
      _operation = new(Context.Tenants, this, NullLogger<DirectoryReadOperation>.Instance);
      Applications = new(_operation);
      Principals = new(_operation);
    }

    public Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(CancellationToken cancellationToken = default)
    {
      cancellationToken.ThrowIfCancellationRequested();
      CreatedForTenants.Add(Context.Tenants.CurrentTenant.Value.TenantId);
      return Task.FromResult(Result.Success<GraphServiceClient, GraphOperationError>(
        new GraphServiceClient(new HttpClient(new Transport(_send)), new AnonymousAuthenticationProvider())));
    }

    public async Task<Result<DirectoryContinuation?, GraphOperationError>> SearchAsync(bool application,
      string? search = null, DirectoryContinuation? next = null)
    {
      if (application)
        return (await Applications.SearchAsync(search, next, TestContext.Current.CancellationToken)).Map(page => page.NextPage);
      return (await Principals.SearchAsync(search, next, TestContext.Current.CancellationToken)).Map(page => page.NextPage);
    }

    public void Dispose() => _operation.Dispose();
  }

  private sealed class Transport(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) : HttpMessageHandler
  {
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
      Assert.Equal(HttpMethod.Get, request.Method);
      return send(request, cancellationToken);
    }
  }
}
