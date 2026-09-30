using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Identity.Web;
using MudBlazor.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"))
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddInMemoryTokenCaches();

builder.Services.AddOptions<MicrosoftIdentityOptions>(OpenIdConnectDefaults.AuthenticationScheme)
    .Validate(options => options.TenantId == "organizations",
        "AzureAd:TenantId must be organizations for organizational multi-tenant sign-in.")
    .Validate(options => Guid.TryParse(options.ClientId, out var id) && id != Guid.Empty,
        "AzureAd:ClientId must be a non-empty application ID.")
    .ValidateOnStart();

builder.Services.AddOptions<CookieAuthenticationOptions>(CookieAuthenticationDefaults.AuthenticationScheme)
    .Configure<ILoggerFactory>((options, loggerFactory) =>
    {
        var previousSignedIn = options.Events.OnSignedIn;
        options.Events.OnSignedIn = async context =>
        {
            await previousSignedIn(context);
            loggerFactory.CreateLogger("MEPA.Authentication").LogInformation("User signed in");
        };
    });

builder.Services.AddAuthorization(options =>
    options.FallbackPolicy = new AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build());
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddControllersWithViews();
builder.Services.AddRazorComponents().AddInteractiveServerComponents();
builder.Services.AddMudServices();
builder.Services.AddAdministrationTenants(builder.Configuration);

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();

app.MapStaticAssets().AllowAnonymous();
app.MapControllers();
app.MapRazorComponents<App>().AddInteractiveServerRenderMode();

app.Run();

public partial class Program;
