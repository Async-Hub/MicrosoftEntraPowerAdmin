# Microsoft Entra Power Admin (MEPA)

.NET 10 Blazor Interactive Server foundation for organizational Microsoft Entra sign-in and configured administration tenant selection.

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
- Each selection resolves against the validated configuration. Changes notify subscribers through a scoped event; UI subscribers unsubscribe on disposal. Future asynchronous tenant-dependent operations must capture the tenant at their start and discard/cancel stale work when selection changes.
- Microsoft.Identity.Web/MSAL owns delegated token acquisition and its server-side in-memory cache. No Graph permission scopes or Graph operations are added yet. Future tenant-specific acquisition must explicitly target the selected tenant and handle consent; changing this selector does not obtain tokens or grant permissions.
- This foundation permits any authenticated organizational user to enter. The administration allowlist limits selectable tenants; it does not establish administrator rights. Role validation is deferred as requested.

Authentication follows the standard [Microsoft.Identity.Web configuration](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-web-app-call-api-app-configuration). UI services follow [MudBlazor setup](https://mudblazor.com/getting-started/installation).

## Verification

Offline tests cover configuration rejection/startup validation, initial tenant rules, switching, invalid selection, notifications, scope isolation, application startup, authentication challenges, protected administration, and sign-out antiforgery enforcement. They use isolated in-memory configuration and local authentication doubles, with no real credentials or Entra/Graph calls.

To verify the real identity-provider round trip after configuration: sign in with an organizational account, check the displayed identity, switch between configured tenants, open a second tab to verify independent selection, and sign out. With several tenants and an unlisted sign-in tenant, confirm that the selector asks for an explicit choice.
