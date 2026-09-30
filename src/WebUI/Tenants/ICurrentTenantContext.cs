using System.Security.Claims;
using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public interface ICurrentTenantContext
{
    IReadOnlyList<MicrosoftEntraTenant> Tenants { get; }
    Maybe<MicrosoftEntraTenant> CurrentTenant { get; }
    event Action? TenantChanged;

    // Called with the server authentication state, once per circuit. Never from browser claims.
    void Initialize(ClaimsPrincipal user);
    Result SelectTenant(Guid tenantId);
}
