using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Authentication;

[Route("account")]
public sealed class AccountController(
    ICurrentTenantContext tenantContext,
    ITokenAcquisition tokenAcquisition) : Controller
{
  [Authorize]
  [HttpPost("connect-tenant")]
  [ValidateAntiForgeryToken]
  public async Task<IActionResult> ConnectTenant(
      [FromForm] Guid tenantId,
      CancellationToken cancellationToken)
  {
    if (tenantContext.SelectTenant(tenantId).IsFailure)
      return BadRequest("The selected tenant is not configured for administration.");

    string? claims = null;
    try
    {
      // Reacquire at this HTTP boundary so claims requirements come from Entra,
      // not browser input or circuit state. No token leaves authentication code.
      _ = await tokenAcquisition.GetAccessTokenForUserAsync(
          [GraphScopes.UserRead, GraphScopes.ApplicationReadAll], authenticationScheme: OpenIdConnectDefaults.AuthenticationScheme,
          tenantId: tenantId.ToString(), user: User,
          tokenAcquisitionOptions: new TokenAcquisitionOptions
          {
            ForceRefresh = true,
            CancellationToken = cancellationToken
          });
    }
    catch (MicrosoftIdentityWebChallengeUserException exception)
    {
      claims = exception.MsalUiRequiredException.Claims;
    }
    catch (MsalUiRequiredException exception)
    {
      claims = exception.Claims;
    }
    catch (MsalException)
    {
      return Problem("Microsoft Entra could not connect to this tenant. Check account access and application configuration.",
          statusCode: StatusCodes.Status403Forbidden);
    }

    var properties = new OpenIdConnectChallengeProperties
    {
      RedirectUri = Url.Content($"~/admin?tenant={tenantId:D}"),
      Scope = ["openid", "profile", "offline_access", GraphScopes.UserRead, GraphScopes.ApplicationReadAll]
    };
    properties.SetParameter("login_hint", User.GetLoginHint());
    properties.Items[TenantAuthentication.TenantProperty] = tenantId.ToString();
    if (!string.IsNullOrWhiteSpace(claims))
      properties.Items["claims"] = claims;

    return Challenge(properties, OpenIdConnectDefaults.AuthenticationScheme);
  }

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
