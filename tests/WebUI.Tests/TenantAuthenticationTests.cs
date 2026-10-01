using System.Security.Claims;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Authentication;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Http;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class TenantAuthenticationTests
{
  private readonly Guid _tenantA = Guid.NewGuid();
  private readonly Guid _tenantB = Guid.NewGuid();
  private readonly AuthenticationScheme _scheme = new(OpenIdConnectDefaults.AuthenticationScheme, null, typeof(OpenIdConnectHandler));

  [Fact]
  public async Task Code_redemption_uses_protected_tenant_and_overwrites_incoming_hint()
  {
    var options = new OpenIdConnectOptions();
    var redeemedTenants = new List<string>();
    options.Events.OnAuthorizationCodeReceived = context =>
    {
      redeemedTenants.Add(context.ProtocolMessage.DomainHint);
      return Task.CompletedTask;
    };
    Configure(options);

    foreach (var tenant in new[] { _tenantA, _tenantB, _tenantA })
    {
      var context = new AuthorizationCodeReceivedContext(new DefaultHttpContext(), _scheme, options, Properties(tenant))
      {
        ProtocolMessage = new OpenIdConnectMessage { DomainHint = "untrusted-browser-hint" }
      };
      await options.Events.AuthorizationCodeReceived(context);
      Assert.Null(context.Result);
    }

    Assert.Equal(new[] { _tenantA, _tenantB, _tenantA }.Select(id => id.ToString()), redeemedTenants);
    Assert.Equal("/signin-oidc", options.CallbackPath);
    Assert.Equal("/signout-callback-oidc", options.SignedOutCallbackPath);
  }

  [Theory]
  [InlineData("invalid")]
  [InlineData("00000000-0000-0000-0000-000000000000")]
  [InlineData("045738d0-b287-4dcf-bbd7-c6a3d32a0b9f")]
  public async Task Invalid_or_removed_tenant_is_rejected_before_code_redemption(string tenant)
  {
    var options = new OpenIdConnectOptions();
    var called = false;
    options.Events.OnAuthorizationCodeReceived = _ =>
    {
      called = true;
      return Task.CompletedTask;
    };
    Configure(options);
    var properties = new AuthenticationProperties();
    properties.Items[TenantAuthentication.TenantProperty] = tenant;
    var context = new AuthorizationCodeReceivedContext(new DefaultHttpContext(), _scheme, options, properties)
    {
      ProtocolMessage = new OpenIdConnectMessage()
    };

    await options.Events.AuthorizationCodeReceived(context);

    Assert.False(called);
    Assert.Contains("not configured", context.Result?.Failure?.Message);
  }

  [Theory]
  [InlineData(true)]
  [InlineData(false)]
  public async Task Returned_identity_must_match_selected_tenant(bool matches)
  {
    var options = new OpenIdConnectOptions();
    var originalHandlerCalled = false;
    options.Events.OnTokenValidated = _ =>
    {
      originalHandlerCalled = true;
      return Task.CompletedTask;
    };
    Configure(options);
    var principal = new ClaimsPrincipal(new ClaimsIdentity(
      [new Claim("tid", (matches ? _tenantA : _tenantB).ToString())], "Test"));
    var context = new TokenValidatedContext(new DefaultHttpContext(), _scheme, options, principal, Properties(_tenantA));

    await options.Events.TokenValidated(context);

    Assert.True(originalHandlerCalled);
    if (matches)
      Assert.Null(context.Result);
    else
      Assert.Contains("different tenant", context.Result?.Failure?.Message);
  }

  [Fact]
  public async Task Normal_login_preserves_organizational_sign_in_and_ignores_incoming_tenant_hint()
  {
    var options = new OpenIdConnectOptions();
    string? redeemedTenant = "not-called";
    options.Events.OnAuthorizationCodeReceived = context =>
    {
      redeemedTenant = context.ProtocolMessage.DomainHint;
      return Task.CompletedTask;
    };
    Configure(options);
    var code = new AuthorizationCodeReceivedContext(new DefaultHttpContext(), _scheme, options, new AuthenticationProperties())
    {
      ProtocolMessage = new OpenIdConnectMessage { DomainHint = _tenantB.ToString() }
    };
    await options.Events.AuthorizationCodeReceived(code);
    Assert.Null(redeemedTenant);

    var principal = new ClaimsPrincipal(new ClaimsIdentity([new Claim("tid", Guid.NewGuid().ToString())], "Test"));
    var validated = new TokenValidatedContext(new DefaultHttpContext(), _scheme, options, principal, new AuthenticationProperties());
    await options.Events.TokenValidated(validated);
    Assert.Null(validated.Result);
  }

  [Fact]
  public async Task Tenant_marker_is_removed_before_the_sign_in_cookie_is_created()
  {
    var options = new OpenIdConnectOptions();
    var originalHandlerCalled = false;
    options.Events.OnTicketReceived = _ =>
    {
      originalHandlerCalled = true;
      return Task.CompletedTask;
    };
    Configure(options);
    var properties = Properties(_tenantA);
    properties.RedirectUri = $"/admin?tenant={_tenantA}";
    var ticket = new AuthenticationTicket(new ClaimsPrincipal(), properties, _scheme.Name);
    var context = new TicketReceivedContext(new DefaultHttpContext(), _scheme, options, ticket);

    await options.Events.TicketReceived(context);

    Assert.True(originalHandlerCalled);
    Assert.False(context.Properties.Items.ContainsKey(TenantAuthentication.TenantProperty));
    Assert.Equal($"/admin?tenant={_tenantA}", context.ReturnUri);
  }

  private void Configure(OpenIdConnectOptions options) => TenantAuthentication.Configure(options, new MicrosoftEntraPowerAdminOptions
  {
    Tenants = [new() { Name = "Tenant A", TenantId = _tenantA }, new() { Name = "Tenant B", TenantId = _tenantB }]
  });

  private static AuthenticationProperties Properties(Guid tenant)
  {
    var properties = new AuthenticationProperties();
    properties.Items[TenantAuthentication.TenantProperty] = tenant.ToString();
    return properties;
  }
}
