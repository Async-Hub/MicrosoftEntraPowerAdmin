# Microsoft Entra Power Admin (MEPA)

.NET 10 Blazor Interactive Server foundation for organizational Microsoft Entra sign-in, configured administration tenant selection, and tenant-aware delegated Microsoft Graph access.

## Local configuration

1. Register a **Web** application in Microsoft Entra ID with supported account type **Accounts in any organizational directory**. Personal Microsoft accounts are not supported.
2. For the existing HTTPS launch profile, register the redirect URIs `https://localhost:7110/signin-oidc` and `https://localhost:7110/signout-callback-oidc`. Keep implicit access-token and ID-token grants disabled; Microsoft.Identity.Web uses the authorization code flow.
3. Create a client credential. For local development, keep the client secret value in user-secrets. Never put a real secret in the checked-in settings files.
4. Configure the application ID and at least one allowed administration tenant. Run from the repository root, replacing the placeholders:

```powershell
dotnet user-secrets set "AzureAd:ClientId" "<application-client-id>" --project src/WebUI
dotnet user-secrets set "AzureAd:ClientSecret" "<client-secret-value>" --project src/WebUI
dotnet user-secrets set "MicrosoftEntraPowerAdmin:Tenants:0:Name" "<tenant-friendly-name>" --project src/WebUI
dotnet user-secrets set "MicrosoftEntraPowerAdmin:Tenants:0:TenantId" "<directory-tenant-id>" --project src/WebUI
```

Add tenant index `1`, `2`, etc. for additional tenants. `AzureAd:TenantId` remains `organizations`; it is the sign-in authority, not the selected administration tenant. The checked-in tenant array and client ID are intentionally empty: startup fails until valid configuration is supplied. Tenant IDs must be non-empty GUIDs and unique. Names must be nonblank and unique ignoring case and surrounding whitespace.

```powershell
dotnet restore MEPA.slnx
dotnet build MEPA.slnx --configuration Release
dotnet test MEPA.slnx --configuration Release
dotnet run --project src/WebUI --launch-profile https
```

Use a trusted local ASP.NET Core HTTPS development certificate. Open `https://localhost:7110`, sign in, and select an administration tenant. Sign out submits an antiforgery-protected POST, clears the application cookie, and signs out through Entra.

## Foundation behavior

- `/` is a public sign-in page. `/admin` requires authentication at the server endpoint and component router. New administrative pages belong under `Components/Pages/Admin`, whose imports apply `[Authorize]`.
- The header shows the user's display name, UPN/email when present, and sign-out action. The layout includes the configured tenant selector.
- `ICurrentTenantContext` is scoped to each Interactive Server circuit. A single tenant is selected after authentication. With multiple tenants, the sign-in `tid` is selected only if allowlisted; otherwise selection stays empty. Sign-in from other organizational tenants remains valid.
- Selection is kept in memory in that circuit. A full reload, new tab, or new circuit initializes again from the authenticated principal. Component navigation and reconnecting to a retained circuit preserve the selection. No tenant state is stored in cookies, browser storage, or static fields.
- Each selection resolves against the validated configuration. Changes notify subscribers through a scoped event; UI subscribers unsubscribe on disposal. The dashboard immediately clears the previous organization, cancels superseded loads, and ignores late results, including rapid A → B → A switches.
- Microsoft.Identity.Web/MSAL owns delegated token acquisition and its server-side in-memory cache. Each Graph operation gets a new official `GraphServiceClient` bound to the selected tenant and the current circuit's authenticated principal. Token acquisition explicitly supplies `tenantId`; using only the sign-in `tid` would target the wrong directory when administering another tenant. Tokens remain inside authentication infrastructure.
- This foundation permits any authenticated organizational user to enter. The administration allowlist limits selectable tenants; it does not establish administrator rights. Role validation is deferred as requested.

Authentication follows the standard [Microsoft.Identity.Web configuration](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-web-app-call-api-app-configuration). UI services follow [MudBlazor setup](https://mudblazor.com/getting-started/installation).

## Microsoft Graph and tenant consent

The administration dashboard reads `id`, `displayName`, and `verifiedDomains` through `GraphServiceClient.Organization.GetAsync` with `$select`, using the public-cloud Microsoft Graph v1.0 endpoint. It verifies the returned organization ID against the selected tenant before displaying a connected state. Empty, ambiguous, or mismatched responses fail validation. No administration operations are implemented.

Add **Microsoft Graph → Delegated permissions → User.Read** to the application registration. This is sufficient for these three properties according to the [organization API documentation](https://learn.microsoft.com/en-us/graph/api/organization-list?view=graph-rest-1.0). `Organization.Read.All` and future administration scopes are not requested. Normal sign-in retains its existing OIDC scopes; Graph consent is requested when needed.

Register just these two Web redirect URIs on the multi-tenant application, shared by every configured tenant:

```text
https://localhost:7110/signin-oidc
https://localhost:7110/signout-callback-oidc
```

Use your application's HTTPS origin instead of localhost when applicable. Previously registered `/signin-oidc/<tenant-guid>` callbacks can be removed after restarting MEPA with this version. Adding an administration tenant no longer requires another redirect URI.

When silent acquisition cannot satisfy consent, MFA, Conditional Access, or an empty token cache, the dashboard displays **Interaction required** with an **Authenticate tenant** action. This submits an authenticated, antiforgery-protected request that validates the tenant against configuration and invokes the Microsoft.Identity.Web/OpenID Connect challenge for that tenant. MSAL claims requirements are obtained again server-side and passed to the challenge. One OIDC handler uses tenant discovery to select the authorization endpoint while retaining the shared callback. The selected tenant travels in middleware-protected authentication state and is supplied to Identity.Web's code-redemption handler. The returned identity must match that tenant; the Graph organization check remains an additional safeguard. Shared authentication options are never changed per request, and no authorization URLs are constructed by MEPA. After authentication, the dashboard restores the selected tenant through an allowlist-validated return parameter and verifies Graph access again.

The built-in Identity.Web Blazor challenge helper does not accept a tenant override. The Graph integration package offers request-level tenant/user overrides; MEPA instead uses a small Kiota access-token adapter to bind both values once when creating each client, preventing callers from accidentally omitting those overrides. Microsoft.Identity.Web still performs all token acquisition and caching.

The signed-in account must be permitted to access the selected directory as a member or guest. Consent is tenant-specific: selecting another configured tenant does not grant access. If that tenant disables user consent, an administrator must grant the delegated permission under its policy. Entra enforces account access, MFA, and Conditional Access. Rejected access appears separately from interaction requirements, throttling, and other Graph failures. A full challenge refreshes the shared authentication cookie, so other tabs may require a reload.

## Verification

Offline tests cover configuration validation, initial tenant rules, circuit isolation, token tenant/user binding, tenant switching, organization mapping and identity verification, stale responses, cancellation, Graph/MSAL errors, tenant-specific challenge scopes and claims, protected administration, and antiforgery enforcement. They use isolated configuration, local authentication doubles, and the real Graph SDK with an in-memory transport, with no real credentials or Entra/Graph calls.

To verify the real identity-provider round trip after configuration: sign in, select Tenant A, complete **Authenticate tenant** if prompted, and confirm its actual organization and connected status. Switch to Tenant B and confirm A's information disappears immediately and B's organization appears after its own authentication/consent. Test a tenant where your account has no access, and verify that a denied request never reports a connected state. Also check independent tab selection and sign-out. With several tenants and an unlisted sign-in tenant, confirm that the selector asks for an explicit choice.
