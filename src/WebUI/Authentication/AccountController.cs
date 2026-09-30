using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Authentication;

[Route("account")]
public sealed class AccountController : Controller
{
    [AllowAnonymous]
    [HttpGet("signin")]
    public IActionResult SignInWithMicrosoftEntra() => Challenge(
        new AuthenticationProperties { RedirectUri = Url.Content("~/admin") },
        OpenIdConnectDefaults.AuthenticationScheme);

    [Authorize]
    [HttpPost("signout")]
    [ValidateAntiForgeryToken]
    public IActionResult SignOutOfMicrosoftEntra() => SignOut(
        new AuthenticationProperties { RedirectUri = Url.Content("~/") },
        CookieAuthenticationDefaults.AuthenticationScheme,
        OpenIdConnectDefaults.AuthenticationScheme);
}
