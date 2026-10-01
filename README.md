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

The administration dashboard reads `id`, `displayName`, and `verifiedDomains` through `GraphServiceClient.Organization.GetAsync` with `$select`, using the public-cloud Microsoft Graph v1.0 endpoint. It verifies the returned organization ID against the selected tenant before displaying a connected state. Empty, ambiguous, or mismatched responses fail validation. Application browsing is read-only.

Add **Microsoft Graph → Delegated permissions → User.Read**, **Application.Read.All**, **Policy.Read.All**, and **Policy.ReadWrite.ApplicationConfiguration** to the MEPA application registration. `User.Read` supports the dashboard's three organization properties according to the [organization API documentation](https://learn.microsoft.com/en-us/graph/api/organization-list?view=graph-rest-1.0). **Application.Read.All, Policy.Read.All, and Policy.ReadWrite.ApplicationConfiguration require admin consent in each administered tenant**. The read permissions enable discovery of applications, service principals, and claims mapping policies with their assignments; the policy write permission enables policy creation. **Application.ReadWrite.All is not requested**. Normal sign-in retains its OIDC scopes; the existing **Authenticate tenant** flow requests the Graph permissions when interaction is needed.

Register just these two Web redirect URIs on the multi-tenant application, shared by every configured tenant:

```text
https://localhost:7110/signin-oidc
https://localhost:7110/signout-callback-oidc
```

Use your application's HTTPS origin instead of localhost when applicable. Previously registered `/signin-oidc/<tenant-guid>` callbacks can be removed after restarting MEPA with this version. Adding an administration tenant no longer requires another redirect URI.

When silent acquisition cannot satisfy consent, MFA, Conditional Access, or an empty token cache, the dashboard displays **Interaction required** with an **Authenticate tenant** action. This submits an authenticated, antiforgery-protected request that validates the tenant against configuration and invokes the Microsoft.Identity.Web/OpenID Connect challenge for that tenant. MSAL claims requirements are obtained again server-side and passed to the challenge. One OIDC handler uses tenant discovery to select the authorization endpoint while retaining the shared callback. The selected tenant travels in middleware-protected authentication state and is supplied to Identity.Web's code-redemption handler. The returned identity must match that tenant; the Graph organization check remains an additional safeguard. Shared authentication options are never changed per request, and no authorization URLs are constructed by MEPA. After authentication, the dashboard restores the selected tenant through an allowlist-validated return parameter and verifies Graph access again.

The built-in Identity.Web Blazor challenge helper does not accept a tenant override. The Graph integration package offers request-level tenant/user overrides; MEPA instead uses a small Kiota access-token adapter to bind both values once when creating each client, preventing callers from accidentally omitting those overrides. Microsoft.Identity.Web still performs all token acquisition and caching.

The signed-in account must be permitted to access the selected directory as a member or guest. Consent is tenant-specific: selecting another configured tenant does not grant access. If that tenant disables user consent, an administrator must grant the delegated permission under its policy. Entra enforces account access, MFA, and Conditional Access. Rejected access appears separately from interaction requirements, throttling, and other Graph failures. A full challenge refreshes the shared authentication cookie, so other tabs may require a reload.

## Read-only application browsing

Under **Applications**, open **App Registrations** or **Enterprise Applications**. Select **Search** to browse, enter a display-name prefix, or enter a GUID to match a Client ID or the corresponding Object ID. Filtering runs on Graph; results use 25-item pages and **Next page** follows Graph's next link. Changing search text requires Search again. No tenant-wide collection is downloaded.

- App Registration = **Application object**. Its **Application Object ID** is `Application.Id`.
- Enterprise Application = **Service Principal object**. Its **Service Principal Object ID** is `ServicePrincipal.Id`.
- **Application / Client ID** (`AppId`) is the shared relationship identifier. The two Object IDs identify different directory objects and are never interchangeable.

Details show copyable IDs, the configured current tenant, read-only properties, and the related object found by Client ID. API settings preserve **True**, **False**, and **Not configured**. Service principals show the verified publisher when available; the v1.0 SDK does not expose a service-principal creation timestamp or a standalone publisher-name property. A local service principal can belong to an App Registration owned by another tenant, so no local App Registration is a normal outcome. An App Registration without a local service principal is also handled normally.

Tenant changes clear results, search text, continuation state, and details immediately. Details return to the corresponding list. Pending operations are canceled, late responses are ignored (including A → B → A), and each new service operation resolves a fresh tenant-bound Graph client. Paging links are retained only on the server and rejected after any tenant change. The existing organization validation on the dashboard is preserved. No editing or Graph write operations are provided.

## Read-only Claims Mapping Policy discovery

Open **Claims Mapping Policies** to load the current tenant's policies. The list shows names, Policy Object IDs, definition availability, and parsed schema/transformation counts when available. Details show `IsOrganizationDefault`, structured definitions, formatted original JSON, and objects returned by `appliesTo`. Each definition string is parsed independently with `System.Text.Json`; missing fields, unknown properties, malformed JSON, and unsupported structures leave the original text available. Both `ClaimsTransformation` and the plural spelling in Graph examples are understood. Unknown fields remain in the raw view. Counts are unavailable if a definition cannot be interpreted completely enough to calculate a total.

Enterprise Application details include an **Assigned Claims Mapping Policies** panel. It matches only typed service principals' `Id` against the selected **Service Principal Object ID**. Application objects returned by `appliesTo` are displayed separately with **Application Object ID** labels; neither that ID nor the shared `AppId` is used to infer service-principal assignments.

All requests use the installed official Microsoft Graph .NET SDK and v1.0: `Policies.ClaimsMappingPolicies.GetAsync`, `Policies.ClaimsMappingPolicies[id].GetAsync`, and `Policies.ClaimsMappingPolicies[id].AppliesTo.GetAsync`. Both collections follow validated next links with the SDK's `WithUrl` builder. Reverse discovery reads each policy's relationship rather than relying on `$expand`: Microsoft documents directory-object expansion limits that can truncate relationships without next links, so expansion cannot reliably establish a complete assignment list. There is no assignment cache or cross-tenant reuse. Tenant changes cancel and clear the old state; the policy list reloads, and policy details return to the list.

The documented least-privilege read permissions are [Policy.Read.All for policies](https://learn.microsoft.com/en-us/graph/api/claimsmappingpolicy-list?view=graph-rest-1.0) and [Policy.Read.All plus Application.Read.All for appliesTo](https://learn.microsoft.com/en-us/graph/api/claimsmappingpolicy-list-appliesto?view=graph-rest-1.0). **ServicePrincipals[id].ClaimsMappingPolicies is never called**. Microsoft also documents a [claims mapping policy permission issue](https://learn.microsoft.com/en-us/graph/known-issues#claims-mapping-policy-might-require-consent-to-additional-permissions) that can cause list/detail reads to return 403 despite the read permissions. Step 6 requests Policy.ReadWrite.ApplicationConfiguration for creation. Real-tenant availability must be checked after consent.

## Create a Claims Mapping Policy

Choose **Create policy** from the policy list. Enter a Display Name, select Include Basic Claim Set, and optionally add up to 50 claims. Each claim can use a directory source (`user`, `application`, `resource`, `audience`, or `company`) with an attribute ID, or a constant Value. Known attribute choices are suggested; other supported IDs can be entered. Supply a JWT claim type, a SAML claim type, or both.

Choose **Review** to validate and see the read-only generated inner JSON. Changing fields requires reviewing again. **Create policy** explicitly submits the reviewed snapshot; controls are disabled while creation is in progress. **Cancel** discards local edits without a Graph request. Tenant switching discards the draft, and the service rejects its original selection identity even after switching back. A request already submitted may complete in its original tenant; check that tenant's list if switching during creation.

Creation uses the official v1.0 SDK's `Policies.ClaimsMappingPolicies.PostAsync` with a `Microsoft.Graph.Models.ClaimsMappingPolicy`, as described in the [create API documentation](https://learn.microsoft.com/en-us/graph/api/claimsmappingpolicy-post-claimsmappingpolicies?view=graph-rest-1.0). The controlled create model is separate from the tolerant discovery model. `System.Text.Json` serializes a typed definition envelope with Version fixed at 1 and IncludeBasicClaimSet encoded as a string Boolean; null optional properties and irrelevant source/value fields are omitted. The resulting JSON string is the single member of the SDK model's `Definition` collection.

The list refreshes after creation and retains the returned Graph resource as authoritative, including when the refresh fails or discovery has not yet picked it up. Its definitions appear read-only with a link to full details. Safe failures retain Graph status, code, and request identifiers for logging. For an uncertain network/timeout outcome, refresh the list before retrying creation.

Cloning, editing, deleting, assigning, unassigning, transformation editing, directory extension discovery, manifest changes, `acceptMappedClaims` configuration, and token testing/generation remain deferred. Newly created policies have no assignments. The raw definition remains a read-only view.

## Verification

Offline tests cover configuration validation, initial tenant rules, circuit isolation, token tenant/user binding, tenant switching, organization mapping and identity verification, stale responses, cancellation, Graph/MSAL errors, tenant-specific challenge scopes and claims, protected administration, and antiforgery enforcement. They use isolated configuration, local authentication doubles, and the real Graph SDK with an in-memory transport, with no real credentials or Entra/Graph calls.

Directory tests additionally cover model mapping, nullable API settings, display-name and GUID searches, exact Client ID lookups, paging and next-link validation, read-only requests, and service/page tenant isolation. Rendered details tests verify both relationship directions, missing related objects, and clearing IDs while returning to the list after tenant selection changes.

Policy tests cover parsing and raw JSON preservation, malformed/partial definitions, policy and relationship paging, safe Graph/MSAL errors, cancellation, service-principal assignment correlation, protected policy routes, rendered structured/raw views, and clearing/reloading policy state on tenant changes. Creation tests cover serialization, validation, the 50-claim limit, safe write errors, nested definition strings, the authoritative POST response, and tenant changes before/during creation. Discovery transports assert GET-only requests; creation transports permit only the policy POST, with no assignment requests.

To verify the real identity-provider round trip after configuration: sign in, select Tenant A, complete **Authenticate tenant** if prompted, and confirm its actual organization and connected status. Switch to Tenant B and confirm A's information disappears immediately and B's organization appears after its own authentication/consent. Test a tenant where your account has no access, and verify that a denied request never reports a connected state. Also check independent tab selection and sign-out. With several tenants and an unlisted sign-in tenant, confirm that the selector asks for an explicit choice.
