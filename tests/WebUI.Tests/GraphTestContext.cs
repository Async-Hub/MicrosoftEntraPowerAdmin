using System.Security.Claims;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

internal sealed class GraphTestContext
{
  public Guid TenantA { get; } = Guid.NewGuid();
  public Guid TenantB { get; } = Guid.NewGuid();
  public CurrentTenantContext Tenants { get; }
  public ClaimsPrincipal User { get; }
  public AuthenticationStateProvider Authentication { get; }

  public GraphTestContext()
  {
    Tenants = new(Options.Create(new MicrosoftEntraPowerAdminOptions
    {
      Tenants = [new() { Name = "Tenant A", TenantId = TenantA }, new() { Name = "Tenant B", TenantId = TenantB }]
    }), NullLogger<CurrentTenantContext>.Instance);
    User = new(new ClaimsIdentity([new Claim("tid", TenantA.ToString()), new Claim("oid", "test-user")], "Test"));
    Authentication = new TestAuthenticationStateProvider(User);
  }

  private sealed class TestAuthenticationStateProvider(ClaimsPrincipal user) : AuthenticationStateProvider
  {
    public override Task<AuthenticationState> GetAuthenticationStateAsync() => Task.FromResult(new AuthenticationState(user));
  }
}
