using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using MudBlazor;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Layout;

public partial class MainLayout(
    AuthenticationStateProvider authenticationStateProvider,
    ICurrentTenantContext tenantContext) : LayoutComponentBase
{
    private readonly MudTheme Theme = new()
    {
        PaletteLight = new PaletteLight
        {
            Primary = "#0078D4",
            PrimaryContrastText = "#FFFFFF",
            Secondary = "#005A9E",
            SecondaryContrastText = "#FFFFFF",
            Background = "#F5F5F5",
            Surface = "#FFFFFF",
            AppbarBackground = "#0078D4",
            AppbarText = "#FFFFFF",
            TextPrimary = "#242424",
            TextSecondary = "#616161"
        }
    };

    protected override async Task OnInitializedAsync()
    {
        var authenticationState = await authenticationStateProvider.GetAuthenticationStateAsync();
        tenantContext.Initialize(authenticationState.User);
    }
}
