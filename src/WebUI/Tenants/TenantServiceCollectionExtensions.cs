using Microsoft.Extensions.Options;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public static class TenantServiceCollectionExtensions
{
    public static IServiceCollection AddAdministrationTenants(
        this IServiceCollection services, IConfiguration configuration)
    {
        services.AddSingleton<IValidateOptions<MicrosoftEntraPowerAdminOptions>, MicrosoftEntraPowerAdminOptionsValidator>();
        services.AddOptions<MicrosoftEntraPowerAdminOptions>()
            .Bind(configuration.GetSection(MicrosoftEntraPowerAdminOptions.SectionName))
            .ValidateOnStart();
        services.AddScoped<ICurrentTenantContext, CurrentTenantContext>();
        return services;
    }
}
