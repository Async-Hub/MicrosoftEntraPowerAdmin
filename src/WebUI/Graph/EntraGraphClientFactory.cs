using System.Security.Claims;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions.Authentication;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed class EntraGraphClientFactory(
  ICurrentTenantContext tenantContext,
  AuthenticationStateProvider authenticationStateProvider,
  ITokenAcquisition tokenAcquisition) : IEntraGraphClientFactory
{
  public async Task<Result<GraphServiceClient, GraphOperationError>> CreateAsync(
    CancellationToken cancellationToken = default)
  {
    cancellationToken.ThrowIfCancellationRequested();
    var tenant = tenantContext.CurrentTenant;
    if (tenant.HasNoValue)
      return new GraphOperationError(GraphOperationErrorType.TenantNotSelected, "Select a tenant to connect to Microsoft Graph.");

    var state = await authenticationStateProvider.GetAuthenticationStateAsync().WaitAsync(cancellationToken);
    if (state.User.Identity?.IsAuthenticated != true)
      return new GraphOperationError(GraphOperationErrorType.Unauthorized, "Sign in to connect to Microsoft Graph.");

    if (tenantContext.CurrentTenant != tenant)
      return new GraphOperationError(GraphOperationErrorType.TenantChanged, "The selected tenant changed. Connect again.");

    // Identity.Web.GraphServiceClient 4.15 supplies tenant/user overrides as request options.
    // This adapter binds both once, so callers cannot omit a request-level override.
    // Kiota owns bearer-header handling; Identity.Web/MSAL owns acquisition and caching.
    var provider = new BaseBearerTokenAuthenticationProvider(
      new TenantAccessTokenProvider(tokenAcquisition, tenant.Value.TenantId, state.User));
    return new GraphServiceClient(provider);
  }

  private sealed class TenantAccessTokenProvider(
    ITokenAcquisition tokenAcquisition,
    Guid tenantId,
    ClaimsPrincipal user) : IAccessTokenProvider
  {
    public AllowedHostsValidator AllowedHostsValidator { get; } = new(["graph.microsoft.com"]);

    public async Task<string> GetAuthorizationTokenAsync(
      Uri uri,
      Dictionary<string, object>? additionalAuthenticationContext = null,
      CancellationToken cancellationToken = default)
    {
      cancellationToken.ThrowIfCancellationRequested();
      if (uri.Scheme != Uri.UriSchemeHttps || !AllowedHostsValidator.IsUrlHostValid(uri))
        return string.Empty;

      return await tokenAcquisition.GetAccessTokenForUserAsync(
        [GraphScopes.UserRead, GraphScopes.ApplicationReadAll],
        authenticationScheme: OpenIdConnectDefaults.AuthenticationScheme,
        tenantId: tenantId.ToString(),
        user: user,
        tokenAcquisitionOptions: new TokenAcquisitionOptions { CancellationToken = cancellationToken });
    }
  }
}
