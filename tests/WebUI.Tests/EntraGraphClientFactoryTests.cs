using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions;
using NSubstitute;
using System.Security.Claims;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class EntraGraphClientFactoryTests
{
  private readonly GraphTestContext _context = new();
  private readonly ITokenAcquisition _tokens = Substitute.For<ITokenAcquisition>();

  [Fact]
  public async Task Missing_selection_returns_failure_without_token_acquisition()
  {
    var result = await CreateFactory().CreateAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantNotSelected, result.Error.Type);
    Assert.Empty(_tokens.ReceivedCalls());
  }

  [Fact]
  public async Task Switching_tenants_creates_distinct_clients_and_uses_selected_tenant_and_current_user()
  {
    var calls = new List<(string? Tenant, ClaimsPrincipal? User, string[] Scopes, CancellationToken Cancellation)>();
    _tokens.GetAccessTokenForUserAsync(Arg.Any<IEnumerable<string>>(), Arg.Any<string>(), Arg.Any<string>(),
      Arg.Any<string>(), Arg.Any<ClaimsPrincipal>(), Arg.Any<TokenAcquisitionOptions>())
      .Returns(call =>
      {
        Assert.Equal(OpenIdConnectDefaults.AuthenticationScheme, call.ArgAt<string>(1));
        calls.Add((call.ArgAt<string>(2), call.ArgAt<ClaimsPrincipal>(4),
          call.ArgAt<IEnumerable<string>>(0).ToArray(), call.ArgAt<TokenAcquisitionOptions>(5).CancellationToken));
        return Task.FromResult("offline-test-token");
      });
    var factory = CreateFactory();
    _context.Tenants.SelectTenant(_context.TenantA);
    var createdA = await factory.CreateAsync(TestContext.Current.CancellationToken);
    Assert.True(createdA.IsSuccess);
    using var clientA = createdA.Value;
    using var requestA = await clientA.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
      clientA.Organization.ToGetRequestInformation(), TestContext.Current.CancellationToken);

    _context.Tenants.SelectTenant(_context.TenantB);
    var createdB = await factory.CreateAsync(TestContext.Current.CancellationToken);
    Assert.True(createdB.IsSuccess);
    using var clientB = createdB.Value;
    using var requestB = await clientB.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
      clientB.Organization.ToGetRequestInformation(), TestContext.Current.CancellationToken);
    // A retained operation remains bound to A; resolving a new operation gives B.
    using var secondRequestA = await clientA.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
      clientA.Organization.ToGetRequestInformation(), TestContext.Current.CancellationToken);

    Assert.NotSame(clientA, clientB);
    Assert.Equal([_context.TenantA.ToString(), _context.TenantB.ToString(), _context.TenantA.ToString()], calls.Select(call => call.Tenant));
    Assert.All(calls, call =>
    {
      Assert.Same(_context.User, call.User);
      Assert.Equal([GraphScopes.UserRead, GraphScopes.ApplicationReadAll, GraphScopes.PolicyReadAll], call.Scopes);
      Assert.Equal(TestContext.Current.CancellationToken, call.Cancellation);
    });
    Assert.Equal("Bearer", requestB?.Headers.Authorization?.Scheme);
  }

  [Fact]
  public async Task Anonymous_user_is_rejected()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var authentication = Substitute.For<AuthenticationStateProvider>();
    authentication.GetAuthenticationStateAsync().Returns(new AuthenticationState(new ClaimsPrincipal()));
    var result = await new EntraGraphClientFactory(_context.Tenants, authentication, _tokens)
      .CreateAsync(TestContext.Current.CancellationToken);
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.Unauthorized, result.Error.Type);
    Assert.Empty(_tokens.ReceivedCalls());
  }

  [Fact]
  public async Task Selection_changed_while_resolving_user_is_rejected()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var pending = new TaskCompletionSource<AuthenticationState>();
    var authentication = Substitute.For<AuthenticationStateProvider>();
    authentication.GetAuthenticationStateAsync().Returns(pending.Task);
    var operation = new EntraGraphClientFactory(_context.Tenants, authentication, _tokens)
      .CreateAsync(TestContext.Current.CancellationToken);
    _context.Tenants.SelectTenant(_context.TenantB);
    pending.SetResult(new AuthenticationState(_context.User));
    var result = await operation;
    Assert.True(result.IsFailure);
    Assert.Equal(GraphOperationErrorType.TenantChanged, result.Error.Type);
  }

  [Theory]
  [InlineData("https://example.com/organization")]
  [InlineData("http://graph.microsoft.com/v1.0/organization")]
  public async Task Token_is_not_acquired_for_an_untrusted_destination(string url)
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    var result = await CreateFactory().CreateAsync(TestContext.Current.CancellationToken);
    using var client = result.Value;
    using var request = await client.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
      new RequestInformation { URI = new Uri(url), HttpMethod = Method.GET }, TestContext.Current.CancellationToken);
    Assert.Null(request?.Headers.Authorization);
    Assert.Empty(_tokens.ReceivedCalls());
  }

  [Fact]
  public async Task Cancellation_is_propagated_without_acquiring_a_token()
  {
    _context.Tenants.SelectTenant(_context.TenantA);
    using var cancellation = new CancellationTokenSource();
    cancellation.Cancel();
    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => CreateFactory().CreateAsync(cancellation.Token));
    Assert.Empty(_tokens.ReceivedCalls());
  }

  private EntraGraphClientFactory CreateFactory() => new(_context.Tenants, _context.Authentication, _tokens);
}
