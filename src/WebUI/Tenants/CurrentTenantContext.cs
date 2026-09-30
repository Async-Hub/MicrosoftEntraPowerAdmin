using System.Security.Claims;
using CSharpFunctionalExtensions;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Web;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public sealed class CurrentTenantContext(
    IOptions<MicrosoftEntraPowerAdminOptions> options,
    ILogger<CurrentTenantContext> logger) : ICurrentTenantContext
{
    private bool _initialized;

    public IReadOnlyList<MicrosoftEntraTenant> Tenants { get; } = Array.AsReadOnly(
        options.Value.Tenants.Select(tenant => new MicrosoftEntraTenant(tenant.TenantId, tenant.Name.Trim())).ToArray());

    public Maybe<MicrosoftEntraTenant> CurrentTenant { get; private set; }

    public event Action? TenantChanged;

    public void Initialize(ClaimsPrincipal user)
    {
        if (_initialized || user.Identity?.IsAuthenticated != true)
            return;

        _initialized = true;
        var initialTenant = Tenants.Count == 1
            ? Maybe<MicrosoftEntraTenant>.From(Tenants[0])
            : Guid.TryParse(user.GetTenantId(), out var tenantId)
                ? FindTenant(tenantId)
                : Maybe<MicrosoftEntraTenant>.None;

        if (CurrentTenant.HasValue || initialTenant.HasNoValue)
            return;

        CurrentTenant = initialTenant;
        logger.LogInformation("Initial administration tenant selected: {TenantId}", initialTenant.Value.TenantId);
        TenantChanged?.Invoke();
    }

    public Result SelectTenant(Guid tenantId)
    {
        var tenant = FindTenant(tenantId);
        if (tenant.HasNoValue)
        {
            logger.LogWarning("Invalid administration tenant selection attempted: {TenantId}", tenantId);
            return Result.Failure("The selected tenant is not configured for administration.");
        }

        if (CurrentTenant.HasValue && CurrentTenant.Value == tenant.Value)
            return Result.Success();

        CurrentTenant = tenant;
        logger.LogInformation("Administration tenant changed to {TenantId}", tenantId);
        TenantChanged?.Invoke();
        return Result.Success();
    }

    private Maybe<MicrosoftEntraTenant> FindTenant(Guid tenantId) =>
        Maybe<MicrosoftEntraTenant>.From(Tenants.FirstOrDefault(tenant => tenant.TenantId == tenantId));
}
