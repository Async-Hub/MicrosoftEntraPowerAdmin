using System.Net;
using System.Text;
using System.Text.Json;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using CSharpFunctionalExtensions;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions.Authentication;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class OrganizationServiceTests
{
  private readonly GraphTestContext _context = new();

  [Fact]
  public async Task Maps_organization_and_requests_only_required_fields()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var service = Service((request, _) =>
    {
      Assert.Equal("https", request.RequestUri?.Scheme);
      Assert.Equal("graph.microsoft.com", request.RequestUri?.Host);
      Assert.Equal("/v1.0/organization", request.RequestUri?.AbsolutePath);
      Assert.Equal("?$select=id,displayName,verifiedDomains", Uri.UnescapeDataString(request.RequestUri?.Query ?? ""));
      return Task.FromResult(Organization(_context.TenantA, "Actual directory", "contoso.test", "CONTOSO.test", " "));
    });

    var result = await service.GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal(_context.TenantA, result.Value.TenantId);
    Assert.Equal("Actual directory", result.Value.DisplayName);
    Assert.Equal(["contoso.test"], result.Value.VerifiedDomains);
  }

  [Fact]
  public async Task Mismatching_organization_returns_failure()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var result = await Service((_, _) => Task.FromResult(Organization(_context.TenantB, "Wrong directory")))
      .GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantMismatch, result.Error.Type);
    Assert.Contains(_context.TenantA.ToString(), result.Error.Message);
    Assert.Contains(_context.TenantB.ToString(), result.Error.Message);
  }

  [Theory]
  [InlineData("{}")]
  [InlineData("{\"value\":[]}")]
  [InlineData("{\"value\":[{},{}]}")]
  [InlineData("{\"value\":[{}]}")]
  [InlineData("{\"value\":[{\"id\":\"invalid\"}]}")]
  [InlineData("{\"value\":[{\"id\":\"00000000-0000-0000-0000-000000000000\"}]}")]
  public async Task Missing_or_ambiguous_organization_returns_meaningful_failure(string json)
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var result = await Service((_, _) => Task.FromResult(Json(json))).GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    Assert.Contains("organization", result.Error.Message);
  }

  [Fact]
  public async Task Null_name_and_domains_are_handled()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var result = await Service((_, _) => Task.FromResult(Json(JsonSerializer.Serialize(new
    {
      value = new[] { new { id = _context.TenantA, displayName = (string?)null } }
    })))).GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsSuccess);
    Assert.Equal("Name unavailable", result.Value.DisplayName);
    Assert.Empty(result.Value.VerifiedDomains);
  }

  [Theory]
  [InlineData(401, GraphOperationErrorType.Unauthorized)]
  [InlineData(403, GraphOperationErrorType.Forbidden)]
  [InlineData(404, GraphOperationErrorType.TenantInaccessible)]
  [InlineData(429, GraphOperationErrorType.Throttled)]
  [InlineData(500, GraphOperationErrorType.GraphFailure)]
  public async Task Maps_Graph_errors_and_preserves_diagnostic_metadata(int status, GraphOperationErrorType expected)
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var service = Service((_, _) => Task.FromResult(Json("""
      {"error":{"code":"TestGraphError","message":"Internal sensitive details",
        "innerError":{"request-id":"request-123","client-request-id":"client-123"}}}
      """, (HttpStatusCode)status)));
    var result = await service.GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(expected, result.Error.Type);
    Assert.Equal(status, result.Error.HttpStatus);
    Assert.Equal("TestGraphError", result.Error.Code);
    Assert.Equal("request-123", result.Error.RequestId);
    Assert.Equal("client-123", result.Error.ClientRequestId);
    Assert.DoesNotContain("Internal sensitive details", result.Error.Message);
  }

  [Theory]
  [InlineData(false, false, GraphOperationErrorType.InteractionRequired)]
  [InlineData(true, false, GraphOperationErrorType.InteractionRequired)]
  [InlineData(true, true, GraphOperationErrorType.ConsentRequired)]
  public async Task Distinguishes_interaction_and_consent(bool wrapped, bool consent, GraphOperationErrorType expected)
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var msal = new MsalUiRequiredException(consent ? "consent_required" : "invalid_grant", "Do not display this");
    Exception exception = wrapped ? new MicrosoftIdentityWebChallengeUserException(msal, [GraphScopes.UserRead]) : msal;
    var result = await Service((_, _) => Task.FromException<HttpResponseMessage>(exception))
      .GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(expected, result.Error.Type);
    Assert.Contains("authentication or consent", result.Error.Message);
  }

  [Fact]
  public async Task No_selection_does_not_create_a_client()
  {
    var factory = new TestClientFactory(() => throw new InvalidOperationException("Must not create a client"));
    var result = await new OrganizationService(_context.Tenants, factory, NullLogger<OrganizationService>.Instance)
      .GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantNotSelected, result.Error.Type);
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Network_failure_or_timeout_returns_a_safe_Graph_error(bool timeout)
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    Exception exception = timeout ? new TaskCanceledException("Transport timeout") : new HttpRequestException("Transport details");
    var result = await Service((_, _) => Task.FromException<HttpResponseMessage>(exception))
      .GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.GraphFailure, result.Error.Type);
    Assert.DoesNotContain("Transport", result.Error.Message);
  }

  [Fact]
  public async Task Switching_A_to_B_returns_organization_B_with_a_new_client()
  {
    var resolvedTenants = new List<Guid>();
    var factory = new TestClientFactory(() =>
    {
      var tenant = _context.Tenants.CurrentTenant.Value;
      resolvedTenants.Add(tenant.TenantId);
      return Client((_, _) => Task.FromResult(Organization(tenant.TenantId, $"Organization {tenant.Name}")));
    });
    var service = new OrganizationService(_context.Tenants, factory, NullLogger<OrganizationService>.Instance);
    _context.Tenants.SelectTenant(_context.TenantA);
    var resultA = await service.GetCurrentAsync(TestContext.Current.CancellationToken);
    _context.Tenants.SelectTenant(_context.TenantB);
    var resultB = await service.GetCurrentAsync(TestContext.Current.CancellationToken);
    Assert.True(resultA.IsSuccess);
    Assert.True(resultB.IsSuccess);
    Assert.Equal(_context.TenantA, resultA.Value.TenantId);
    Assert.Equal(_context.TenantB, resultB.Value.TenantId);
    Assert.Equal("Organization Tenant B", resultB.Value.DisplayName);
    Assert.Equal([_context.TenantA, _context.TenantB], resolvedTenants);
  }

  [Fact]
  public async Task Late_response_for_A_is_rejected_after_switching_to_B()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var pending = new TaskCompletionSource<HttpResponseMessage>();
    var operation = Service((_, _) => pending.Task).GetCurrentAsync(TestContext.Current.CancellationToken);
    _context.Tenants.SelectTenant(_context.TenantB);
    pending.SetResult(Organization(_context.TenantA, "Stale A"));
    var result = await operation;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
  }

  [Fact]
  public async Task Cancellation_reaches_the_SDK_and_is_not_mapped_to_Graph_failure()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    using var cancellation = new CancellationTokenSource();
    var service = Service(async (_, token) =>
    {
      cancellation.Cancel();
      await Task.Delay(Timeout.Infinite, token);
      throw new InvalidOperationException("Unreachable");
    });
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => service.GetCurrentAsync(cancellation.Token));
  }

  private OrganizationService Service(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) =>
    new(_context.Tenants, new TestClientFactory(() => Client(send)), NullLogger<OrganizationService>.Instance);

  private static GraphServiceClient Client(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) =>
    new(new HttpClient(new TestTransport(send)), new AnonymousAuthenticationProvider());

  private static HttpResponseMessage Organization(Guid id, string name, params string[] domains) => Json(JsonSerializer.Serialize(new
  {
    value = new[] { new { id, displayName = name, verifiedDomains = domains.Select(domain => new { name = domain }) } }
  }));

  private static HttpResponseMessage Json(string json, HttpStatusCode status = HttpStatusCode.OK) => new(status)
  {
    Content = new StringContent(json, Encoding.UTF8, "application/json")
  };

  private sealed class TestTransport(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) : HttpMessageHandler
  {
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
      send(request, cancellationToken);
  }

  private sealed class TestClientFactory(Func<GraphServiceClient> create) : IEntraGraphClientFactory
  {
    public Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(CancellationToken cancellationToken = default) =>
      Task.FromResult(Result.Success<GraphServiceClient, GraphOperationError>(create()));
  }
}
