using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class TenantOptionsTests
{
    [Fact]
    public void Valid_configuration_is_accepted()
    {
        var result = Validate(new MicrosoftEntraTenantOptions { Name = "Tenant A", TenantId = Guid.NewGuid() });
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void Empty_collection_is_rejected() => AssertFailure(Validate(), "At least one");

    [Fact]
    public void Empty_tenant_id_is_rejected() =>
        AssertFailure(Validate(new MicrosoftEntraTenantOptions { Name = "Tenant A" }), "IDs must not be empty");

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Empty_name_is_rejected(string name) =>
        AssertFailure(Validate(new MicrosoftEntraTenantOptions { Name = name, TenantId = Guid.NewGuid() }), "names must not be empty");

    [Fact]
    public void Duplicate_ids_are_rejected()
    {
        var id = Guid.NewGuid();
        AssertFailure(Validate(
            new() { Name = "Tenant A", TenantId = id },
            new() { Name = "Tenant B", TenantId = id }), "IDs must be unique");
    }

    [Theory]
    [InlineData("Tenant A")]
    [InlineData(" tenant a ")]
    public void Ambiguous_names_are_rejected(string name) => AssertFailure(Validate(
        new() { Name = "Tenant A", TenantId = Guid.NewGuid() },
        new() { Name = name, TenantId = Guid.NewGuid() }), "names must be unique");

    [Fact]
    public async Task Invalid_configuration_fails_at_host_start()
    {
        var builder = new HostApplicationBuilder(new HostApplicationBuilderSettings
        {
            DisableDefaults = true
        });
        builder.Services.AddAdministrationTenants(new ConfigurationBuilder().Build());
        using var host = builder.Build();

        var exception = await Assert.ThrowsAsync<OptionsValidationException>(() => host.StartAsync(TestContext.Current.CancellationToken));
        Assert.Contains("At least one", exception.Message);
    }

    private static ValidateOptionsResult Validate(params MicrosoftEntraTenantOptions[] tenants) =>
        new MicrosoftEntraPowerAdminOptionsValidator().Validate(null, new() { Tenants = tenants });

    private static void AssertFailure(ValidateOptionsResult result, string message)
    {
        Assert.True(result.Failed);
        Assert.Contains(message, result.FailureMessage);
    }
}
