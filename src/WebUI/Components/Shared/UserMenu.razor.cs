using System.Security.Claims;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;

public partial class UserMenu
{
    private static string GetDisplayName(ClaimsPrincipal user) =>
        user.FindFirst("name")?.Value ?? user.Identity?.Name ?? "Signed-in user";

    private static string? GetUserName(ClaimsPrincipal user) =>
        user.FindFirst("preferred_username")?.Value
        ?? user.FindFirst(ClaimTypes.Upn)?.Value
        ?? user.FindFirst("upn")?.Value
        ?? user.FindFirst(ClaimTypes.Email)?.Value
        ?? user.FindFirst("email")?.Value;
}
