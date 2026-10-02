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

Add **Microsoft Graph → Delegated permissions → User.Read**, **Application.Read.All**, **Policy.Read.All**, **Policy.ReadWrite.ApplicationConfiguration**, and **Application.ReadWrite.All** to the MEPA application registration. `User.Read` supports the dashboard's three organization properties according to the [organization API documentation](https://learn.microsoft.com/en-us/graph/api/organization-list?view=graph-rest-1.0). **All four Application/Policy permissions require admin consent in each administered tenant**. The read permissions enable discovery; the policy write permission enables policy creation; Step 7 adds Application.ReadWrite.All for direct assignment discovery and assignment/unassignment, together with policy read access already satisfied by Policy.ReadWrite.ApplicationConfiguration. Normal sign-in retains its OIDC scopes; the existing **Authenticate tenant** flow requests the Graph permissions when interaction is needed.

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

Enterprise Application details include an **Assigned Claims Mapping Policies** panel, now read directly through `ServicePrincipals[id].ClaimsMappingPolicies.GetAsync` using the **Service Principal Object ID**. Policy details continue to page `appliesTo`; application objects returned there are displayed separately with **Application Object ID** labels. Neither that ID nor the shared `AppId` is used for assignment mutations.

All requests use the installed official Microsoft Graph .NET SDK and v1.0. Collections follow validated next links with the SDK's `WithUrl` builder. The Step 5 policy-side reverse discovery remains available, without relying on `$expand`, while Step 7's Service Principal view uses the direct relationship. There is no assignment cache or cross-tenant reuse. Tenant changes cancel and clear the old state; the policy list reloads, and policy details return to the list.

The documented least-privilege read permissions are [Policy.Read.All for policies](https://learn.microsoft.com/en-us/graph/api/claimsmappingpolicy-list?view=graph-rest-1.0) and [Policy.Read.All plus Application.Read.All for appliesTo](https://learn.microsoft.com/en-us/graph/api/claimsmappingpolicy-list-appliesto?view=graph-rest-1.0). [Direct assigned-policy discovery](https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-claimsmappingpolicies?view=graph-rest-1.0) requires Application.ReadWrite.All. Microsoft also documents a [claims mapping policy permission issue](https://learn.microsoft.com/en-us/graph/known-issues#claims-mapping-policy-might-require-consent-to-additional-permissions) that can cause list/detail reads to return 403 despite the read permissions. Real-tenant availability must be checked after consent.

## Create a Claims Mapping Policy

Choose **Create policy** from the policy list. Enter a Display Name, select Include Basic Claim Set, and optionally add up to 50 claims. Each claim can use a directory source (`user`, `application`, `resource`, `audience`, or `company`) with an attribute ID, or a constant Value. Known attribute choices are suggested; other supported IDs can be entered. Supply a JWT claim type, a SAML claim type, or both.

Choose **Review** to validate and see the read-only generated inner JSON. Changing fields requires reviewing again. **Create policy** explicitly submits the reviewed snapshot; controls are disabled while creation is in progress. **Cancel** discards local edits without a Graph request. Tenant switching discards the draft, and the service rejects its original selection identity even after switching back. A request already submitted may complete in its original tenant; check that tenant's list if switching during creation.

Creation uses the official v1.0 SDK's `Policies.ClaimsMappingPolicies.PostAsync` with a `Microsoft.Graph.Models.ClaimsMappingPolicy`, as described in the [create API documentation](https://learn.microsoft.com/en-us/graph/api/claimsmappingpolicy-post-claimsmappingpolicies?view=graph-rest-1.0). The controlled create model is separate from the tolerant discovery model. `System.Text.Json` serializes a typed definition envelope with Version fixed at 1 and IncludeBasicClaimSet encoded as a string Boolean; null optional properties and irrelevant source/value fields are omitted. The resulting JSON string is the single member of the SDK model's `Definition` collection.

The list refreshes after creation and retains the returned Graph resource as authoritative, including when the refresh fails or discovery has not yet picked it up. Its definitions appear read-only with a link to full details. Safe failures retain Graph status, code, and request identifiers for logging. For an uncertain network/timeout outcome, refresh the list before retrying creation.

Cloning, editing, transformation editing, directory extension discovery, manifest changes, `acceptMappedClaims` configuration, and token testing/generation remain deferred. Newly created policies have no assignments. The raw definition remains a read-only view.

## Assign and unassign Claims Mapping Policies

From Enterprise Application details, choose **Assign policy**, filter existing policies by name or Policy Object ID, and select a policy. Already assigned policies are excluded. From policy details, choose **Assign to Enterprise Application**, use the existing paged Enterprise Application search, and select a Service Principal. IDs are displayed for clarity; neither workflow requires pasting IDs. Existing relationships are shown and remain in place; Graph determines whether another assignment is accepted.

The shared confirmation identifies the policy, Enterprise Application, Service Principal Object ID, and tenant. Before **Assign**, it explains that Claims Mapping Policy takes precedence over Custom Claims Policy and claims configured through the Entra admin center, as described in [Microsoft's claims customization documentation](https://learn.microsoft.com/en-us/entra/identity-platform/reference-claims-customization). **Unassign** requires a separate confirmation and removes only the relationship, preserving all three directory objects. Opening the confirmation or changing selections never mutates Graph.

Assignment uses `ServicePrincipals[servicePrincipalObjectId].ClaimsMappingPolicies.Ref.PostAsync` with `ReferenceCreate.OdataId` pointing to the v1.0 policy URL ([assignment API](https://learn.microsoft.com/en-us/graph/api/serviceprincipal-post-claimsmappingpolicies?view=graph-rest-1.0)). Unassignment uses `ClaimsMappingPolicies[policyObjectId].Ref.DeleteAsync` ([unassignment API](https://learn.microsoft.com/en-us/graph/api/serviceprincipal-delete-claimsmappingpolicies?view=graph-rest-1.0)). Both revalidate the two objects and current relationship on one tenant-bound client before mutation. Every attempt closes the confirmation and reloads Graph state in the active view, including policy-side `appliesTo`. Successful mutations show a notification; refresh errors remain visible independently.

Busy controls prevent duplicate submissions. Tenant changes discard selections and invalidate the confirmation identity even after switching back. Already submitted operations may complete in the original tenant; check that tenant's assignments before retrying. Safe errors distinguish missing consent/privileges, missing objects, duplicate/stale relationships, throttling, and uncertain network/timeout results without exposing authentication details.

## Delete a Claims Mapping Policy

Choose **Delete policy** in policy details. MEPA retrieves the latest policy and its paged `appliesTo` relationships before opening a dedicated confirmation. Any assigned directory object blocks deletion. Review the displayed Enterprise Applications and use their existing **Open** or **Unassign** actions; each unassignment remains a separate, explicitly confirmed operation. Reopen deletion after removing the final relationship.

The confirmation identifies the current Display Name, Policy Object ID, and tenant, shows the available definition summary and assignment count, and explains that deletion is permanent. Type the exact current policy name (surrounding whitespace is trimmed) to enable the final **Delete policy** button. Missing, whitespace-only, or control-character names use the fixed phrase **DELETE**. Unsupported and malformed definitions can still be deleted without modifying their contents.

Deletion uses the official v1.0 SDK's `Policies.ClaimsMappingPolicies[policyObjectId].DeleteAsync` and the existing delegated **Policy.ReadWrite.ApplicationConfiguration** permission ([delete API documentation](https://learn.microsoft.com/en-us/graph/api/claimsmappingpolicy-delete?view=graph-rest-1.0)). No new permissions are added. The service rechecks the current name and all assignments on one tenant-bound client immediately before DELETE. Busy controls prevent duplicate deletion and simultaneous mutations. Tenant switching cancels and discards confirmation, including switching away and back.

After `204 No Content`, details close and the policy list reloads from Graph. A missing policy, including DELETE returning 404, also returns to the refreshed list with an informative notification. Permission, throttling, and uncertain transport failures use structured errors and refresh details without assuming deletion succeeded. MEPA adds no destructive retry or automatic unassignment. Bulk deletion, backup/restore, import/export, and other cleanup operations remain out of scope.

## Verification

Offline tests cover configuration validation, initial tenant rules, circuit isolation, token tenant/user binding, tenant switching, organization mapping and identity verification, stale responses, cancellation, Graph/MSAL errors, tenant-specific challenge scopes and claims, protected administration, and antiforgery enforcement. They use isolated configuration, local authentication doubles, and the real Graph SDK with an in-memory transport, with no real credentials or Entra/Graph calls.

Directory tests additionally cover model mapping, nullable API settings, display-name and GUID searches, exact Client ID lookups, paging and next-link validation, read-only requests, and service/page tenant isolation. Rendered details tests verify both relationship directions, missing related objects, and clearing IDs while returning to the list after tenant selection changes.

Policy tests cover parsing and raw JSON preservation, malformed/partial definitions, policy and relationship paging, safe Graph/MSAL errors, cancellation, service-principal assignment correlation, protected policy routes, rendered structured/raw views, and clearing/reloading policy state on tenant changes. Creation tests cover serialization, validation, the 50-claim limit, safe write errors, nested definition strings, the authoritative POST response, and tenant changes before/during creation. Discovery transports assert GET-only requests; creation transports permit only the policy POST, with no assignment requests.

Assignment tests use the real SDK with an in-memory transport to verify Object ID routing, policy reference construction, relationship-only DELETE, direct relationship paging, stale state, structured failures, cancellation, and tenant changes at every asynchronous stage. Offline component interaction tests cover both selectors, already assigned filtering, explicit confirmation, and duplicate-submission/tenant-switch guards.

Deletion tests verify zero/one/multiple assignments, newly added relationships and later pages, exact-name and fallback confirmation, renamed/missing policies, malformed definitions, `204` handling, safe Graph/MSAL/transport failures, cancellation, and tenant isolation. Component tests cover fresh confirmation state, reuse of unassignment, busy guards, stale-assignment refresh, list reload and notifications, and disposal during tenant switching.

To verify the real identity-provider round trip after configuration: sign in, select Tenant A, complete **Authenticate tenant** if prompted, and confirm its actual organization and connected status. Switch to Tenant B and confirm A's information disappears immediately and B's organization appears after its own authentication/consent. Test a tenant where your account has no access, and verify that a denied request never reports a connected state. Also check independent tab selection and sign-out. With several tenants and an unlisted sign-in tenant, confirm that the selector asks for an explicit choice.
