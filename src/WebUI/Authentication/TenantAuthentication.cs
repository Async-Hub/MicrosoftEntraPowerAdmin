using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.Identity.Web;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Authentication;

public static class TenantAuthentication
{
  public const string TenantProperty = "mepa:authentication-tenant";

  public static void Configure(OpenIdConnectOptions options, MicrosoftEntraPowerAdminOptions tenants)
  {
    var previousRedirect = options.Events.OnRedirectToIdentityProvider;
    options.Events.OnRedirectToIdentityProvider = async context =>
    {
      await previousRedirect(context);
      if (context.Handled)
        return;

      var selected = SelectedTenant(context.Properties, tenants);
      if (selected.IsFailure)
      {
        context.Response.StatusCode = StatusCodes.Status400BadRequest;
        context.HandleResponse();
        return;
      }
      if (selected.Value.HasNoValue)
        return;

      // Discover the tenant's endpoint; the OIDC middleware builds the authorization
      // request, correlation/nonce cookies and protected state with the shared callback.
      // Never mutate Options.Authority/Configuration: those options are shared by users.
      var metadataAddress = context.Options.MetadataAddress;
      if (string.IsNullOrEmpty(metadataAddress) || !metadataAddress.Contains("/organizations/", StringComparison.Ordinal))
        throw new InvalidOperationException("Tenant authentication requires organizations discovery metadata.");

      var configuration = await OpenIdConnectConfigurationRetriever.GetAsync(
        metadataAddress.Replace("/organizations/", $"/{selected.Value.Value:D}/", StringComparison.Ordinal),
        new HttpDocumentRetriever(context.Options.Backchannel) { RequireHttps = true },
        context.HttpContext.RequestAborted);
      context.ProtocolMessage.IssuerAddress = configuration.AuthorizationEndpoint;
    };

    var previousCodeReceived = options.Events.OnAuthorizationCodeReceived;
    options.Events.OnAuthorizationCodeReceived = async context =>
    {
      var selected = SelectedTenant(context.Properties, tenants);
      if (selected.IsFailure)
      {
        context.Fail(selected.Error);
        return;
      }

      // Identity.Web 4.15's code-redemption handler passes ProtocolMessage.DomainHint
      // as AuthCodeRedemptionParameters.Tenant to MSAL's WithTenantId. Overwrite any
      // incoming hint using only middleware-protected state, before invoking it.
      // See TokenAcquisition-AspnetCore.cs in the microsoft-identity-web 4.15.0 source.
      context.ProtocolMessage.DomainHint = selected.Value.HasValue ? selected.Value.Value.ToString() : null;
      await previousCodeReceived(context);
    };

    var previousTokenValidated = options.Events.OnTokenValidated;
    options.Events.OnTokenValidated = async context =>
    {
      await previousTokenValidated(context);
      if (context.Result is not null)
        return;

      var selected = SelectedTenant(context.Properties, tenants);
      if (selected.IsFailure)
      {
        context.Fail(selected.Error);
        return;
      }
      if (selected.Value.HasValue &&
        (!Guid.TryParse(context.Principal?.GetTenantId(), out var actualTenantId) || actualTenantId != selected.Value.Value))
        context.Fail("Microsoft Entra returned a different tenant than the one selected for authentication.");
    };

    var previousTicketReceived = options.Events.OnTicketReceived;
    options.Events.OnTicketReceived = async context =>
    {
      await previousTicketReceived(context);
      // The tenant marker belongs to this authentication round trip, not the sign-in cookie.
      context.Properties.Items.Remove(TenantProperty);
    };
  }

  private static Result<Maybe<Guid>> SelectedTenant(AuthenticationProperties? properties, MicrosoftEntraPowerAdminOptions tenants)
  {
    if (properties is null || !properties.Items.TryGetValue(TenantProperty, out var value))
      return Result.Success(Maybe<Guid>.None);

    if (!Guid.TryParse(value, out var tenantId) || !tenants.Tenants.Any(tenant => tenant.TenantId == tenantId))
      return Result.Failure<Maybe<Guid>>("The authentication tenant is not configured for administration.");

    return Result.Success(Maybe<Guid>.From(tenantId));
  }
}
