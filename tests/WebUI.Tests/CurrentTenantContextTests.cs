using System.Security.Claims;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class CurrentTenantContextTests
{
    private readonly MicrosoftEntraTenantOptions _tenantA = new() { Name = "Tenant A", TenantId = Guid.NewGuid() };
    private readonly MicrosoftEntraTenantOptions _tenantB = new() { Name = "Tenant B", TenantId = Guid.NewGuid() };

    [Fact]
    public void Context_starts_without_a_selection() => Assert.True(CreateContext().CurrentTenant.HasNoValue);

    [Fact]
    public void Configured_tenant_can_be_selected()
    {
        var context = CreateContext();
        var result = context.SelectTenant(_tenantA.TenantId);
        Assert.True(result.IsSuccess);
        Assert.Equal(new MicrosoftEntraTenant(_tenantA.TenantId, _tenantA.Name), context.CurrentTenant.Value);
    }

    [Fact]
    public void Unknown_tenant_is_rejected_without_creating_a_selection()
    {
        var context = CreateContext();
        var result = context.SelectTenant(Guid.NewGuid());
        Assert.True(result.IsFailure);
        Assert.Contains("not configured", result.Error);
        Assert.True(context.CurrentTenant.HasNoValue);
    }

    [Fact]
    public void Switching_from_A_to_B_updates_selection_and_notifies_subscribers()
    {
        var context = CreateContext();
        Assert.True(context.SelectTenant(_tenantA.TenantId).IsSuccess);
        var notifications = new List<Guid>();
        context.TenantChanged += () => notifications.Add(context.CurrentTenant.Value.TenantId);

        Assert.True(context.SelectTenant(_tenantB.TenantId).IsSuccess);

        Assert.Equal(_tenantB.TenantId, context.CurrentTenant.Value.TenantId);
        Assert.Equal([_tenantB.TenantId], notifications);
    }

    [Fact]
    public void Invalid_selection_preserves_current_tenant_and_does_not_notify()
    {
        var context = CreateContext();
        context.SelectTenant(_tenantA.TenantId);
        var notifications = 0;
        context.TenantChanged += () => notifications++;

        var result = context.SelectTenant(Guid.Empty);

        Assert.True(result.IsFailure);
        Assert.Contains("not configured", result.Error);
        Assert.Equal(_tenantA.TenantId, context.CurrentTenant.Value.TenantId);
        Assert.Equal(0, notifications);
    }

    [Fact]
    public void Selecting_current_tenant_does_not_notify()
    {
        var context = CreateContext();
        context.SelectTenant(_tenantA.TenantId);
        var notifications = 0;
        context.TenantChanged += () => notifications++;
        Assert.True(context.SelectTenant(_tenantA.TenantId).IsSuccess);
        Assert.Equal(0, notifications);
    }

    [Fact]
    public void Single_tenant_is_selected_for_authenticated_user_from_another_tenant()
    {
        var context = CreateContext(_tenantA);
        context.Initialize(User(Guid.NewGuid().ToString()));
        Assert.Equal(_tenantA.TenantId, context.CurrentTenant.Value.TenantId);
    }

    [Theory]
    [InlineData("tid")]
    [InlineData("http://schemas.microsoft.com/identity/claims/tenantid")]
    public void Matching_sign_in_tenant_is_selected(string claimType)
    {
        var context = CreateContext();
        context.Initialize(new ClaimsPrincipal(new ClaimsIdentity(
            [new Claim(claimType, _tenantB.TenantId.ToString())], "Test")));
        Assert.Equal(_tenantB.TenantId, context.CurrentTenant.Value.TenantId);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("invalid")]
    [InlineData("194a53b3-4bb0-443a-8e3d-296c980e93fb")]
    public void Multiple_tenants_without_a_matching_claim_have_no_initial_selection(string? tid)
    {
        var context = CreateContext();
        context.Initialize(User(tid));
        Assert.True(context.CurrentTenant.HasNoValue);
    }

    [Fact]
    public void Anonymous_initialization_does_not_select_a_tenant_or_block_authenticated_initialization()
    {
        var context = CreateContext(_tenantA);
        context.Initialize(new ClaimsPrincipal(new ClaimsIdentity()));
        Assert.True(context.CurrentTenant.HasNoValue);
        context.Initialize(User(null));
        Assert.Equal(_tenantA.TenantId, context.CurrentTenant.Value.TenantId);
    }

    [Fact]
    public void Reinitialization_does_not_overwrite_explicit_selection()
    {
        var context = CreateContext();
        var user = User(_tenantA.TenantId.ToString());
        context.Initialize(user);
        context.SelectTenant(_tenantB.TenantId);
        context.Initialize(user);
        Assert.Equal(_tenantB.TenantId, context.CurrentTenant.Value.TenantId);
    }

    [Fact]
    public void Dependency_injection_scopes_have_independent_tenant_state()
    {
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["MicrosoftEntraPowerAdmin:Tenants:0:Name"] = _tenantA.Name,
            ["MicrosoftEntraPowerAdmin:Tenants:0:TenantId"] = _tenantA.TenantId.ToString(),
            ["MicrosoftEntraPowerAdmin:Tenants:1:Name"] = _tenantB.Name,
            ["MicrosoftEntraPowerAdmin:Tenants:1:TenantId"] = _tenantB.TenantId.ToString()
        }).Build();
        using var provider = new ServiceCollection().AddLogging().AddAdministrationTenants(configuration)
            .BuildServiceProvider(new ServiceProviderOptions { ValidateScopes = true });
        using var scopeA = provider.CreateScope();
        using var scopeB = provider.CreateScope();
        var contextA = scopeA.ServiceProvider.GetRequiredService<ICurrentTenantContext>();
        var contextB = scopeB.ServiceProvider.GetRequiredService<ICurrentTenantContext>();

        contextA.SelectTenant(_tenantA.TenantId);
        Assert.True(contextB.CurrentTenant.HasNoValue);
        contextB.SelectTenant(_tenantB.TenantId);
        Assert.Equal(_tenantA.TenantId, contextA.CurrentTenant.Value.TenantId);
        Assert.Equal(_tenantB.TenantId, contextB.CurrentTenant.Value.TenantId);
        Assert.Same(contextA, scopeA.ServiceProvider.GetRequiredService<ICurrentTenantContext>());
    }

    private CurrentTenantContext CreateContext(params MicrosoftEntraTenantOptions[] tenants) => new(
        Options.Create(new MicrosoftEntraPowerAdminOptions { Tenants = tenants.Length == 0 ? [_tenantA, _tenantB] : tenants }),
        NullLogger<CurrentTenantContext>.Instance);

    private static ClaimsPrincipal User(string? tid) => new(new ClaimsIdentity(
        tid is null ? [] : [new Claim("tid", tid)], "Test"));
}
