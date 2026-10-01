using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.JSInterop;
using MudBlazor.Services;
using NSubstitute;
using ApplicationPage = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin.ApplicationDetails;
using PrincipalPage = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin.ServicePrincipalDetails;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class DirectoryDetailsPageTests
{
  [Theory]
  [InlineData(true, "invalid-id")]
  [InlineData(false, "invalid-id")]
  [InlineData(true, "00000000-0000-0000-0000-000000000000")]
  [InlineData(false, "00000000-0000-0000-0000-000000000000")]
  public async Task Invalid_route_ids_are_rejected_without_service_calls(bool application, string objectId)
  {
    var context = new GraphTestContext();
    context.Tenants.SelectTenant(context.TenantA);
    var applications = Substitute.For<IApplicationService>();
    var principals = Substitute.For<IServicePrincipalService>();
    var services = new ServiceCollection();
    services.AddLogging();
    services.AddMudServices();
    services.AddSingleton<IJSRuntime>(Substitute.For<IJSRuntime>());
    services.AddSingleton<NavigationManager>(new TestNavigation());
    services.AddSingleton<ICurrentTenantContext>(context.Tenants);
    services.AddSingleton(applications);
    services.AddSingleton(principals);
    await using var provider = services.BuildServiceProvider();
    await using var renderer = new HtmlRenderer(provider, provider.GetRequiredService<ILoggerFactory>());
    await renderer.Dispatcher.InvokeAsync(async () =>
    {
      var parameters = ParameterView.FromDictionary(new Dictionary<string, object?> { ["ObjectId"] = objectId });
      var root = application
        ? await renderer.RenderComponentAsync<ApplicationPage>(parameters)
        : await renderer.RenderComponentAsync<PrincipalPage>(parameters);
      Assert.Contains("must be a non-empty GUID", root.ToHtmlString());
    });
    Assert.Empty(applications.ReceivedCalls());
    Assert.Empty(principals.ReceivedCalls());
  }

  [Theory]
  [InlineData(true, true)]
  [InlineData(true, false)]
  [InlineData(false, true)]
  [InlineData(false, false)]
  public async Task Details_resolve_relationship_by_client_id_and_clear_on_tenant_switch(bool application, bool relatedExists)
  {
    var context = new GraphTestContext();
    context.Tenants.SelectTenant(context.TenantA);
    var appObjectId = Guid.NewGuid();
    var principalObjectId = Guid.NewGuid();
    var clientId = Guid.NewGuid();
    var applications = Substitute.For<IApplicationService>();
    var principals = Substitute.For<IServicePrincipalService>();
    var app = new ApplicationDetails(new(appObjectId, clientId, "Local registration", "AzureADMultipleOrgs", null, null),
      null, [], [], [], [], [], null, null);
    var principal = new ServicePrincipalDetails(new(principalObjectId, clientId, "Enterprise instance", "Application", true, null),
      null, null, null, [], context.TenantA, []);
    applications.GetByObjectIdAsync(appObjectId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<ApplicationDetails, GraphOperationError>(app));
    principals.GetByObjectIdAsync(principalObjectId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<ServicePrincipalDetails, GraphOperationError>(principal));
    applications.FindByAppIdAsync(clientId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<Maybe<ApplicationDetails>, GraphOperationError>(relatedExists ? app : Maybe<ApplicationDetails>.None));
    principals.FindByAppIdAsync(clientId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<Maybe<ServicePrincipalDetails>, GraphOperationError>(relatedExists ? principal : Maybe<ServicePrincipalDetails>.None));
    var navigation = new TestNavigation();
    var services = new ServiceCollection();
    services.AddLogging();
    services.AddMudServices();
    services.AddSingleton<IJSRuntime>(Substitute.For<IJSRuntime>());
    services.AddSingleton<NavigationManager>(navigation);
    services.AddSingleton<ICurrentTenantContext>(context.Tenants);
    services.AddSingleton(applications);
    services.AddSingleton(principals);
    await using var provider = services.BuildServiceProvider();
    await using var renderer = new HtmlRenderer(provider, provider.GetRequiredService<ILoggerFactory>());

    await renderer.Dispatcher.InvokeAsync(async () =>
    {
      var parameters = ParameterView.FromDictionary(new Dictionary<string, object?>
      {
        ["ObjectId"] = (application ? appObjectId : principalObjectId).ToString()
      });
      var root = application
        ? await renderer.RenderComponentAsync<ApplicationPage>(parameters)
        : await renderer.RenderComponentAsync<PrincipalPage>(parameters);
      var html = root.ToHtmlString();
      Assert.Contains("Current Tenant: Tenant A", html);
      Assert.Contains(clientId.ToString(), html);
      Assert.Contains((application ? appObjectId : principalObjectId).ToString(), html);
      if (relatedExists)
      {
        Assert.Contains(appObjectId.ToString(), html);
        Assert.Contains(principalObjectId.ToString(), html);
        Assert.Contains(application ? "Open Enterprise Application" : "Open App Registration", html);
      }
      else
        Assert.Contains(application ? "No Enterprise Application / Service Principal exists" : "owned by another tenant", html);
      if (application)
      {
        Assert.Contains("api.acceptMappedClaims: Not configured", html);
        Assert.Contains("api.requestedAccessTokenVersion: Not configured", html);
      }

      context.Tenants.SelectTenant(context.TenantB);
      var cleared = root.ToHtmlString();
      Assert.DoesNotContain(appObjectId.ToString(), cleared);
      Assert.DoesNotContain(principalObjectId.ToString(), cleared);
      Assert.DoesNotContain(clientId.ToString(), cleared);
      Assert.EndsWith(application ? "/applications" : "/service-principals", navigation.Uri);
    });

    if (application)
    {
      await principals.Received(1).FindByAppIdAsync(clientId, Arg.Any<CancellationToken>());
      await principals.DidNotReceive().GetByObjectIdAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>());
    }
    else
    {
      await applications.Received(1).FindByAppIdAsync(clientId, Arg.Any<CancellationToken>());
      await applications.DidNotReceive().GetByObjectIdAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>());
    }
  }

  private sealed class TestNavigation : NavigationManager
  {
    public TestNavigation() => Initialize("https://localhost/", "https://localhost/");
    protected override void NavigateToCore(string uri, bool forceLoad)
    {
      Uri = ToAbsoluteUri(uri).ToString();
      NotifyLocationChanged(false);
    }
  }
}
