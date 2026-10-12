# Microsoft Entra Power Admin (MEPA)

**Making Microsoft Entra administration easier.**

MEPA is a free, open-source web application designed to simplify advanced Microsoft Entra management.

Some Entra features are difficult to access or aren't available directly in the Microsoft Entra admin center. MEPA brings these capabilities together in one convenient interface.

## ✨ What can you do?

- **Manage multiple tenants** — Sign in once and easily switch between your Microsoft Entra tenants.
- **Explore applications** — Browse App Registrations and Enterprise Applications, including their settings and relationships.
- **Manage claims mapping policies** — View, create, assign, unassign, and delete policies without writing scripts or making API calls.

## 🎯 Who is it for?

Microsoft Entra administrators, IT professionals, and cloud engineers who want a simpler way to manage advanced Entra features.

## 🚀 Project Status

MEPA is an evolving open-source project. More features and improvements are planned.

**Contributions, suggestions, and feedback are always welcome!**

## 🚀 Getting Started

### Step 1: Register an Application in Microsoft Entra

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Navigate to **Entra ID → App registrations → New registration**.
3. Enter a name, such as `Microsoft Entra Power Admin`.
4. Under **Supported account types**, select **Accounts in any organizational directory (Multitenant)**.
5. Under **Redirect URI**, select **Web** and enter:
   `https://localhost:7110/signin-oidc`
6. Click **Register**.

#### Configure Authentication

1. Navigate to **Authentication**.
2. Add another Web redirect URI:
   `https://localhost:7110/signout-callback-oidc`
3. Leave implicit grant options disabled.

#### Create a Client Secret

1. Navigate to **Certificates & secrets → Client secrets**.
2. Click **New client secret**.
3. Create a secret and copy its **Value**.

> **Important:** Save the **Application (Client) ID** and **Client Secret**. You'll need them in the next step.

### Step 2: Configure MEPA

Open `src/WebUI/appsettings.json` and update the configuration with your Microsoft Entra application details.

```json
{
  "AzureAd": {
    "Instance": "https://login.microsoftonline.com/",
    "TenantId": "organizations",
    "ClientId": "<your-client-id>",
    "ClientSecret": "",
    "CallbackPath": "/signin-oidc",
    "SignedOutCallbackPath": "/signout-callback-oidc"
  },
  "MicrosoftEntraPowerAdmin": {
    "Tenants": [
      {
        "Name": "My Organization",
        "TenantId": "<tenant-id>"
      },
      {
        "Name": "Another Organization",
        "TenantId": "<tenant-id>"
      }
    ]
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Warning"
    }
  },
  "AllowedHosts": "*"
}
```

**Configuration details:**

* **ClientId** — Application (Client) ID from Step 1.
* **ClientSecret** — Secret created during application registration.
* **Tenants** — Organizations you want to manage. Each requires a friendly name and its Directory (Tenant) ID.

You can add as many tenants as needed or configure just one.

Keep `AzureAd:TenantId` set to `organizations` to allow sign-in with accounts from different Microsoft Entra tenants.

#### Configure the Client Secret

For security, store the client secret using .NET User Secrets rather than saving it in `appsettings.json`:

```powershell
dotnet user-secrets set "AzureAd:ClientSecret" "<your-client-secret>" --project src/WebUI
```

> **Note:** Your account must have access to each configured tenant and sufficient permissions to perform administrative operations.
