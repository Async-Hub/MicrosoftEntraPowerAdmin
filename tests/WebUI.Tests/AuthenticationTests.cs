using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using System.Net;
using System.Security.Claims;
using System.Text.Encodings.Web;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Authentication;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using NSubstitute;


namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class AuthenticationTests
{
  private const string TenantA = "a10fba49-127b-4830-8a31-317e989391af";
  private const string TenantB = "714ed1b1-aebd-47b3-b333-f5221c752d41";

  [Theory]
  [InlineData(TenantA)]
  [InlineData(TenantB)]
  public async Task Tenant_challenge_targets_selected_tenant_and_preserves_required_claims(string tenantId)
  {
    await using var application = new TestApplication(substituteTokens: true);
    using var client = CreateClient(application);
    client.DefaultRequestHeaders.Add("Test-Authenticated", "true");
    var form = await ChallengeFormAsync(application, client, tenantId);

    var response = await client.PostAsync("/account/connect-tenant", new FormUrlEncodedContent(form), TestContext.Current.CancellationToken);

    Assert.Equal(HttpStatusCode.Redirect, response.StatusCode);
    var location = Assert.IsType<Uri>(response.Headers.Location);
    Assert.Equal($"/{tenantId}/oauth2/v2.0/authorize", location.AbsolutePath);
    var query = Microsoft.AspNetCore.WebUtilities.QueryHelpers.ParseQuery(location.Query);
    Assert.Equal("https://localhost/signin-oidc", query["redirect_uri"].ToString());
    Assert.Contains(GraphScopes.UserRead, query["scope"].ToString());
    Assert.DoesNotContain("Application.ReadWrite", query["scope"].ToString());
    Assert.Equal(TestApplication.RequiredClaims, query["claims"].ToString());

    var scheme = OpenIdConnectDefaults.AuthenticationScheme;
    Assert.Equal("organizations", application.Services.GetRequiredService<IOptionsMonitor<MicrosoftIdentityOptions>>().Get(scheme).TenantId);
    var options = application.Services.GetRequiredService<IOptionsMonitor<OpenIdConnectOptions>>().Get(scheme);
    Assert.Equal("Cookies", options.SignInScheme);
    var properties = options.StateDataFormat.Unprotect(query["state"].ToString());
    Assert.Equal($"/admin?tenant={tenantId}", properties?.RedirectUri);
    Assert.Equal(tenantId, properties?.Items[TenantAuthentication.TenantProperty]);
    var schemes = await application.Services.GetRequiredService<IAuthenticationSchemeProvider>().GetAllSchemesAsync();
    Assert.Single(schemes, registered => registered.HandlerType == typeof(OpenIdConnectHandler));

    await application.Tokens.Received().GetAccessTokenForUserAsync(
        Arg.Is<IEnumerable<string>>(scopes => scopes.SequenceEqual(new[] { GraphScopes.UserRead })),
        OpenIdConnectDefaults.AuthenticationScheme, tenantId, null,
        Arg.Is<ClaimsPrincipal>(user => user.Identity != null && user.Identity.IsAuthenticated),
        Arg.Is<TokenAcquisitionOptions>(options => options.ForceRefresh));
  }

  [Fact]
  public async Task Tenant_challenge_rejects_unknown_tenant_before_acquiring_tokens()
  {
    await using var application = new TestApplication(substituteTokens: true);
    using var client = CreateClient(application);
    client.DefaultRequestHeaders.Add("Test-Authenticated", "true");
    var form = await ChallengeFormAsync(application, client, Guid.NewGuid().ToString());
    var response = await client.PostAsync("/account/connect-tenant", new FormUrlEncodedContent(form), TestContext.Current.CancellationToken);
    Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    Assert.Empty(application.Tokens.ReceivedCalls());
  }

  [Fact]
  public async Task Tenant_challenge_requires_antiforgery()
  {
    await using var application = new TestApplication(substituteTokens: true);
    using var client = CreateClient(application);
    client.DefaultRequestHeaders.Add("Test-Authenticated", "true");
    var response = await client.PostAsync("/account/connect-tenant",
        new FormUrlEncodedContent(new Dictionary<string, string> { ["tenantId"] = TenantA }), TestContext.Current.CancellationToken);
    Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    Assert.Empty(application.Tokens.ReceivedCalls());
  }

  private static async Task<Dictionary<string, string>> ChallengeFormAsync(TestApplication application, HttpClient client, string tenantId)
  {
    await using var scope = application.Services.CreateAsyncScope();
    var context = new DefaultHttpContext { RequestServices = scope.ServiceProvider, User = TestUser() };
    context.Request.Scheme = "https";
    var tokens = scope.ServiceProvider.GetRequiredService<IAntiforgery>().GetAndStoreTokens(context);
    client.DefaultRequestHeaders.Add("Cookie", context.Response.Headers.SetCookie.ToString().Split(';')[0]);
    return new()
    {
      [tokens.FormFieldName] = Assert.IsType<string>(tokens.RequestToken),
      ["tenantId"] = tenantId
    };
  }

  [Fact]
  public async Task Public_home_starts_without_authentication()
  {
    await using var application = new TestApplication();
    using var client = CreateClient(application);
    var response = await client.GetAsync("/", TestContext.Current.CancellationToken);
    Assert.Equal(HttpStatusCode.OK, response.StatusCode);
    var html = await response.Content.ReadAsStringAsync(TestContext.Current.CancellationToken);
    Assert.Contains("Sign in", html);
    Assert.Contains("_framework/blazor.", html);
  }

  [Theory]
  [InlineData("/admin")]
  [InlineData("/account/signin")]
  public async Task Anonymous_administration_and_sign_in_challenge_organizational_Entra(string path)
  {
    await using var application = new TestApplication();
    using var client = CreateClient(application);
    var response = await client.GetAsync(path, TestContext.Current.CancellationToken);
    Assert.Equal(HttpStatusCode.Redirect, response.StatusCode);
    var location = Assert.IsType<Uri>(response.Headers.Location);
    Assert.Equal("login.microsoftonline.com", location.Host);
    Assert.Equal("/organizations/oauth2/v2.0/authorize", location.AbsolutePath);
    Assert.Contains("response_type=code", location.Query);
    var query = Microsoft.AspNetCore.WebUtilities.QueryHelpers.ParseQuery(location.Query);
    Assert.Equal("https://localhost/signin-oidc", query["redirect_uri"].ToString());
  }

  [Fact]
  public async Task Authenticated_user_can_enter_administration()
  {
    await using var application = new TestApplication();
    using var client = CreateClient(application);
    client.DefaultRequestHeaders.Add("Test-Authenticated", "true");
    var response = await client.GetAsync("/admin", TestContext.Current.CancellationToken);
    Assert.Equal(HttpStatusCode.OK, response.StatusCode);
  }

  [Fact]
  public async Task Sign_out_requires_an_antiforgery_token()
  {
    await using var application = new TestApplication();
    using var client = CreateClient(application);
    client.DefaultRequestHeaders.Add("Test-Authenticated", "true");
    var response = await client.PostAsync("/account/signout", new FormUrlEncodedContent([]), TestContext.Current.CancellationToken);
    Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
  }

  [Fact]
  public async Task Authenticated_sign_out_clears_cookie_and_redirects_to_Entra()
  {
    await using var application = new TestApplication();
    using var client = CreateClient(application);
    client.DefaultRequestHeaders.Add("Test-Authenticated", "true");
    await using var scope = application.Services.CreateAsyncScope();
    var context = new DefaultHttpContext
    {
      RequestServices = scope.ServiceProvider,
      User = TestUser()
    };
    context.Request.Scheme = "https";
    var tokens = scope.ServiceProvider.GetRequiredService<IAntiforgery>().GetAndStoreTokens(context);
    client.DefaultRequestHeaders.Add("Cookie", context.Response.Headers.SetCookie.ToString().Split(';')[0]);

    var response = await client.PostAsync("/account/signout", new FormUrlEncodedContent(
        new Dictionary<string, string> { [tokens.FormFieldName] = Assert.IsType<string>(tokens.RequestToken) }),
        TestContext.Current.CancellationToken);

    Assert.Equal(HttpStatusCode.Redirect, response.StatusCode);
    Assert.Equal("/organizations/oauth2/v2.0/logout", response.Headers.Location?.AbsolutePath);
    var query = Microsoft.AspNetCore.WebUtilities.QueryHelpers.ParseQuery(response.Headers.Location?.Query ?? "");
    Assert.Equal("https://localhost/signout-callback-oidc", query["post_logout_redirect_uri"].ToString());
    Assert.Contains(response.Headers.GetValues("Set-Cookie"), cookie =>
        cookie.StartsWith(".AspNetCore.Cookies=;", StringComparison.Ordinal));
  }

  [Fact]
  public async Task Anonymous_users_cannot_open_a_Blazor_circuit()
  {
    await using var application = new TestApplication();
    using var client = CreateClient(application);
    var response = await client.PostAsync("/_blazor/negotiate?negotiateVersion=1", null,
        TestContext.Current.CancellationToken);
    Assert.Equal(HttpStatusCode.Redirect, response.StatusCode);
    Assert.Equal("login.microsoftonline.com", response.Headers.Location?.Host);
  }

  private static ClaimsPrincipal TestUser() => new(new ClaimsIdentity(
      [new Claim(ClaimTypes.NameIdentifier, "test-user")], "Test"));

  private static HttpClient CreateClient(TestApplication application) => application.CreateClient(new()
  {
    BaseAddress = new Uri("https://localhost"),
    AllowAutoRedirect = false
  });

  private sealed class TestApplication(bool substituteTokens = false) : WebApplicationFactory<Program>
  {
    public const string RequiredClaims = "{\"access_token\":{\"acrs\":{\"essential\":true,\"value\":\"c1\"}}}";
    public ITokenAcquisition Tokens { get; } = Substitute.For<ITokenAcquisition>();

    private static MsalUiRequiredException ClaimsChallenge()
    {
      var json = System.Text.Json.Nodes.JsonNode.Parse(new MsalUiRequiredException("invalid_grant", "Interaction required").ToJsonString());
      var payload = Assert.IsType<System.Text.Json.Nodes.JsonObject>(json);
      payload["claims"] = RequiredClaims;
      return Assert.IsType<MsalUiRequiredException>(MsalException.FromJsonString(payload.ToJsonString()));
    }

    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
      builder.UseEnvironment("Testing");
      // Isolated host configuration never changes process-wide environment variables.
      foreach (var setting in new Dictionary<string, string>
      {
        ["AzureAd:ClientId"] = "abddf6d2-2cb8-4576-bc48-8bc404037bf7",
        ["MicrosoftEntraPowerAdmin:Tenants:0:Name"] = "Test Tenant",
        ["MicrosoftEntraPowerAdmin:Tenants:0:TenantId"] = TenantA,
        ["MicrosoftEntraPowerAdmin:Tenants:1:Name"] = "Other Test Tenant",
        ["MicrosoftEntraPowerAdmin:Tenants:1:TenantId"] = TenantB
      })
        builder.UseSetting(setting.Key, setting.Value);
      builder.ConfigureTestServices(services =>
      {
        Tokens.GetAccessTokenForUserAsync(Arg.Any<IEnumerable<string>>(), Arg.Any<string>(), Arg.Any<string>(),
                  Arg.Any<string>(), Arg.Any<ClaimsPrincipal>(), Arg.Any<TokenAcquisitionOptions>())
                  .Returns(_ => Task.FromException<string>(new MicrosoftIdentityWebChallengeUserException(
                      ClaimsChallenge(),
                      [GraphScopes.UserRead])));
        if (substituteTokens)
          services.Replace(ServiceDescriptor.Singleton(Tokens));
        services.AddAuthentication(options => options.DefaultAuthenticateScheme = "Test")
                  .AddScheme<AuthenticationSchemeOptions, TestAuthenticationHandler>("Test", _ => { });
        // Static protocol metadata prevents network discovery; no Entra credentials are used.
        services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, options =>
        {
          options.Configuration = new OpenIdConnectConfiguration
          {
            AuthorizationEndpoint = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize",
            EndSessionEndpoint = "https://login.microsoftonline.com/organizations/oauth2/v2.0/logout"
          };
          options.MetadataAddress = "https://login.microsoftonline.com/organizations/v2.0/.well-known/openid-configuration";
          options.Backchannel = new HttpClient(new TenantDiscoveryTransport());
        });
      });
    }
  }

  private sealed class TestAuthenticationHandler(
      IOptionsMonitor<AuthenticationSchemeOptions> options,
      ILoggerFactory logger,
      UrlEncoder encoder) : AuthenticationHandler<AuthenticationSchemeOptions>(options, logger, encoder)
  {
    protected override Task<AuthenticateResult> HandleAuthenticateAsync()
    {
      if (!Request.Headers.ContainsKey("Test-Authenticated"))
        return Task.FromResult(AuthenticateResult.NoResult());

      var principal = TestUser();
      return Task.FromResult(AuthenticateResult.Success(new AuthenticationTicket(principal, Scheme.Name)));
    }
  }

  private sealed class TenantDiscoveryTransport : HttpMessageHandler
  {
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
      var uri = Assert.IsType<Uri>(request.RequestUri);
      Assert.Equal("https", uri.Scheme);
      Assert.Equal("login.microsoftonline.com", uri.Host);
      var tenant = uri.AbsolutePath.Split('/')[1];
      Assert.Contains(tenant, new[] { TenantA, TenantB });
      Assert.Equal($"/{tenant}/v2.0/.well-known/openid-configuration", uri.AbsolutePath);
      return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
      {
        Content = new StringContent(System.Text.Json.JsonSerializer.Serialize(new
        {
          issuer = $"https://login.microsoftonline.com/{tenant}/v2.0",
          authorization_endpoint = $"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"
        }), System.Text.Encoding.UTF8, "application/json")
      });
    }
  }
}
