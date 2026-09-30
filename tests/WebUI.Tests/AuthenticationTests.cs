using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using System.Net;
using System.Security.Claims;
using System.Text.Encodings.Web;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class AuthenticationTests
{
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

    private sealed class TestApplication : WebApplicationFactory<Program>
    {
        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.UseEnvironment("Testing");
            builder.ConfigureAppConfiguration((_, configuration) =>
            {
                configuration.AddInMemoryCollection(new Dictionary<string, string?>
                {
                    ["AzureAd:ClientId"] = "abddf6d2-2cb8-4576-bc48-8bc404037bf7",
                    ["MicrosoftEntraPowerAdmin:Tenants:0:Name"] = "Test Tenant",
                    ["MicrosoftEntraPowerAdmin:Tenants:0:TenantId"] = "a10fba49-127b-4830-8a31-317e989391af"
                });
            });
            builder.ConfigureTestServices(services =>
            {
                services.AddAuthentication(options => options.DefaultAuthenticateScheme = "Test")
                    .AddScheme<AuthenticationSchemeOptions, TestAuthenticationHandler>("Test", _ => { });
                // Static protocol metadata prevents network discovery; no Entra credentials are used.
                services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, options =>
                    options.Configuration = new OpenIdConnectConfiguration
                    {
                        AuthorizationEndpoint = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize",
                        EndSessionEndpoint = "https://login.microsoftonline.com/organizations/oauth2/v2.0/logout"
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
}
