using System.Text.Encodings.Web;
using System.Text.RegularExpressions;
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
    services.AddSingleton(Substitute.For<IClaimsMappingPolicyService>());
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
    var policies = Substitute.For<IClaimsMappingPolicyService>();
    policies.BeginAssignment().Returns(Result.Success<ClaimsMappingPolicyAssignmentContext, GraphOperationError>(new(context.TenantA, Guid.NewGuid())));
    policies.GetAssignedPoliciesAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), principalObjectId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([]));
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
    services.AddSingleton(policies);
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
        Assert.Contains("<dt>Accept mapped claims</dt>", html);
        Assert.Contains("<dt>Requested access token version</dt>", html);
        Assert.Contains("<span class=\"application-detail-empty\">Not configured</span>", html);
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
      await policies.Received(1).GetAssignedPoliciesAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), principalObjectId, Arg.Any<CancellationToken>());
      await applications.Received(1).FindByAppIdAsync(clientId, Arg.Any<CancellationToken>());
      await applications.DidNotReceive().GetByObjectIdAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>());
    }
  }

  [Theory]
  [InlineData(true, "Yes", 2)]
  [InlineData(false, "No", 1)]
  public async Task Application_details_group_and_format_values_without_changing_data(
    bool acceptMappedClaims, string expectedBoolean, int tokenVersion)
  {
    var objectId = Guid.Parse("1188d339-33aa-4594-992b-bac7b3f8aa48");
    var clientId = Guid.Parse("36b18dd0-0138-5599-8c7b-1dd905efbe30");
    var longUri = "https://example.com/" + new string('a', 240) + "?first=1&second=2";
    const string description = "<b>API description</b>";
    var details = new ApplicationDetails(
      new(objectId, clientId, "Api1 <test>", "AzureADMyOrg",
        new DateTimeOffset(2024, 8, 29, 19, 13, 39, TimeSpan.FromHours(4)), "example.onmicrosoft.com"),
      description, ["api://" + clientId, longUri],
      ["https://example.com/web", longUri], ["https://example.com/spa", longUri],
      ["http://localhost", longUri], ["internal", "<script>alert('tag')</script>"],
      acceptMappedClaims, tokenVersion);

    var html = await RenderApplicationAsync(details);

    Assert.Matches("<h1[^>]*>Api1 &lt;test&gt;</h1>", html);
    foreach (var heading in new[] { "General", "API", "Authentication and Identifiers" })
      Assert.Matches($"<h2[^>]*>{heading}</h2>", html);
    Assert.Contains(clientId.ToString(), DetailValue(html, "Application / Client ID"));
    Assert.Contains("Copy Application / Client ID", DetailValue(html, "Application / Client ID"));
    Assert.Contains(objectId.ToString(), DetailValue(html, "Application Object ID"));
    Assert.Contains("Copy Application Object ID", DetailValue(html, "Application Object ID"));
    Assert.Contains(HtmlEncoder.Default.Encode(description), DetailValue(html, "Description"));
    Assert.Contains("AzureADMyOrg", DetailValue(html, "Sign-in Audience"));
    Assert.Contains("Aug 29, 2024, 15:13:39 UTC", DetailValue(html, "Created"));
    Assert.Contains("<span class=\"directory-id\">example.onmicrosoft.com</span>", DetailValue(html, "Publisher Domain"));
    Assert.Contains($">{expectedBoolean}</span>", DetailValue(html, "Accept mapped claims"));
    Assert.Contains($">{tokenVersion}</span>", DetailValue(html, "Requested access token version"));
    Assert.DoesNotContain("api.acceptMappedClaims", html);
    Assert.DoesNotContain("api.requestedAccessTokenVersion", html);
    Assert.DoesNotContain("Not configured", html);

    foreach (var (label, values, technical) in new[]
    {
      ("Web redirect URIs", details.WebRedirectUris, true),
      ("SPA redirect URIs", details.SpaRedirectUris, true),
      ("Public client redirect URIs", details.PublicClientRedirectUris, true),
      ("Identifier URIs", details.IdentifierUris, true),
      ("Tags", details.Tags, false)
    })
    {
      var rendered = DetailValue(html, label);
      Assert.Contains("<ul", rendered);
      Assert.Equal(values.Count, Regex.Matches(rendered, "<li>").Count);
      foreach (var value in values)
        Assert.Contains(HtmlEncoder.Default.Encode(value), rendered);
      Assert.Equal(technical ? values.Count : 0, Regex.Matches(rendered, "class=\"directory-id\"").Count);
    }
    Assert.DoesNotContain("<b>", html);
    Assert.DoesNotContain("<script>", html);
  }

  [Theory]
  [InlineData(null)]
  [InlineData("")]
  [InlineData("   ")]
  public async Task Application_details_render_missing_and_empty_values_consistently(string? emptyValue)
  {
    var details = new ApplicationDetails(
      new(Guid.Parse("1188d339-33aa-4594-992b-bac7b3f8aa48"), null, "Api1", emptyValue, null, emptyValue),
      emptyValue, [], [], [], [], [], null, null);

    var html = await RenderApplicationAsync(details);

    foreach (var label in new[]
    {
      "Application / Client ID", "Description", "Sign-in Audience", "Created", "Publisher Domain",
      "Accept mapped claims", "Requested access token version", "Web redirect URIs", "SPA redirect URIs",
      "Public client redirect URIs", "Identifier URIs", "Tags"
    })
      Assert.Contains("<span class=\"application-detail-empty\">Not configured</span>", DetailValue(html, label));
    Assert.DoesNotContain("Not available", html);
    Assert.DoesNotContain("<ul", html);
    Assert.DoesNotContain("Copy Application / Client ID", html);
  }

  [Theory]
  [InlineData(true, "Yes")]
  [InlineData(false, "No")]
  public async Task Enterprise_application_details_group_and_format_values_without_changing_data(
    bool accountEnabled, string expectedBoolean)
  {
    var objectId = Guid.Parse("1188d339-33aa-4594-992b-bac7b3f8aa48");
    var clientId = Guid.Parse("36b18dd0-0138-5599-8c7b-1dd905efbe30");
    var ownerId = Guid.Parse("a7d6fdf7-f1e2-4a5a-b686-8f4b93eb7d8c");
    var longUri = "https://example.com/" + new string('a', 240) + "?first=1&second=2";
    const string description = "<b>Enterprise description</b>";
    var details = new ServicePrincipalDetails(
      new(objectId, clientId, "Enterprise <test>", "Application", accountEnabled, "Example Publisher"),
      description, longUri, "https://example.com/login?first=1&second=2",
      ["api://" + clientId, longUri], ownerId, ["internal", "<script>alert('tag')</script>"]);

    var html = await RenderPrincipalAsync(details);

    Assert.Matches("<h1[^>]*>Enterprise &lt;test&gt;</h1>", html);
    foreach (var heading in new[] { "General", "Authentication and Identifiers" })
      Assert.Matches($"<h2[^>]*>{heading}</h2>", html);
    foreach (var (label, value) in new[]
    {
      ("Application / Client ID", clientId),
      ("Service Principal Object ID", objectId),
      ("App Owner Organization ID", ownerId)
    })
    {
      Assert.Contains(value.ToString(), DetailValue(html, label));
      Assert.Contains($"Copy {label}", DetailValue(html, label));
    }
    Assert.Contains(HtmlEncoder.Default.Encode(description), DetailValue(html, "Description"));
    Assert.Contains(">Application</span>", DetailValue(html, "Service Principal Type"));
    Assert.Contains($">{expectedBoolean}</span>", DetailValue(html, "Account Enabled"));
    Assert.Contains("Example Publisher", DetailValue(html, "Verified Publisher Name"));
    Assert.Contains($"<span class=\"directory-id\">{HtmlEncoder.Default.Encode(longUri)}</span>", DetailValue(html, "Homepage"));
    Assert.Contains("class=\"directory-id\"", DetailValue(html, "Login URL"));
    Assert.Contains(HtmlEncoder.Default.Encode("https://example.com/login?first=1&second=2"), DetailValue(html, "Login URL"));
    foreach (var (label, values, technical) in new[]
    {
      ("Service Principal Names", details.ServicePrincipalNames, true),
      ("Tags", details.Tags, false)
    })
    {
      var rendered = DetailValue(html, label);
      Assert.Contains("<ul", rendered);
      Assert.Equal(values.Count, Regex.Matches(rendered, "<li>").Count);
      foreach (var value in values)
        Assert.Contains(HtmlEncoder.Default.Encode(value), rendered);
      Assert.Equal(technical ? values.Count : 0, Regex.Matches(rendered, "class=\"directory-id\"").Count);
    }
    Assert.Contains("Assigned Claims Mapping Policies", html);
    Assert.Contains("Refresh policies", html);
    Assert.Contains("Assign policy", html);
    Assert.Contains("No Claims Mapping Policies are assigned to this Service Principal.", html);
    Assert.DoesNotContain("<b>", html);
    Assert.DoesNotContain("<script>", html);
    Assert.DoesNotContain("Not configured", html);
  }

  [Theory]
  [InlineData(null)]
  [InlineData("")]
  [InlineData("   ")]
  public async Task Enterprise_application_details_render_missing_and_empty_values_consistently(string? emptyValue)
  {
    var details = new ServicePrincipalDetails(
      new(Guid.Parse("1188d339-33aa-4594-992b-bac7b3f8aa48"), null, "Enterprise", emptyValue, null, emptyValue),
      emptyValue, emptyValue, emptyValue, [], null, []);

    var html = await RenderPrincipalAsync(details);

    foreach (var label in new[]
    {
      "Application / Client ID", "Description", "Service Principal Type", "Account Enabled",
      "Verified Publisher Name", "App Owner Organization ID", "Homepage", "Login URL",
      "Service Principal Names", "Tags"
    })
      Assert.Contains("<span class=\"application-detail-empty\">Not configured</span>", DetailValue(html, label));
    Assert.DoesNotContain("Not available", html);
    Assert.DoesNotContain("<ul", html);
    Assert.DoesNotContain("Copy Application / Client ID", html);
    Assert.DoesNotContain("Copy App Owner Organization ID", html);
  }

  private static string DetailValue(string html, string label)
  {
    var match = Regex.Match(html, $"<dt>{Regex.Escape(label)}</dt>\\s*<dd>(.*?)</dd>", RegexOptions.Singleline);
    Assert.True(match.Success, $"Missing detail: {label}");
    return match.Groups[1].Value;
  }

  private static async Task<string> RenderApplicationAsync(ApplicationDetails details)
  {
    var context = new GraphTestContext();
    context.Tenants.SelectTenant(context.TenantA);
    var applications = Substitute.For<IApplicationService>();
    var principals = Substitute.For<IServicePrincipalService>();
    applications.GetByObjectIdAsync(details.Application.ObjectId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<ApplicationDetails, GraphOperationError>(details));
    if (details.Application.AppId is { } clientId)
      principals.FindByAppIdAsync(clientId, Arg.Any<CancellationToken>())
        .Returns(Result.Success<Maybe<ServicePrincipalDetails>, GraphOperationError>(Maybe<ServicePrincipalDetails>.None));
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
    return await renderer.Dispatcher.InvokeAsync(async () =>
    {
      var parameters = ParameterView.FromDictionary(new Dictionary<string, object?>
      {
        ["ObjectId"] = details.Application.ObjectId.ToString()
      });
      var root = await renderer.RenderComponentAsync<ApplicationPage>(parameters);
      return root.ToHtmlString();
    });
  }

  private static async Task<string> RenderPrincipalAsync(ServicePrincipalDetails details)
  {
    var context = new GraphTestContext();
    context.Tenants.SelectTenant(context.TenantA);
    var applications = Substitute.For<IApplicationService>();
    var principals = Substitute.For<IServicePrincipalService>();
    var policies = Substitute.For<IClaimsMappingPolicyService>();
    policies.BeginAssignment().Returns(Result.Success<ClaimsMappingPolicyAssignmentContext, GraphOperationError>(
      new(context.TenantA, Guid.Parse("a7d6fdf7-f1e2-4a5a-b686-8f4b93eb7d8c"))));
    policies.GetAssignedPoliciesAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(),
        details.ServicePrincipal.ObjectId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([]));
    principals.GetByObjectIdAsync(details.ServicePrincipal.ObjectId, Arg.Any<CancellationToken>())
      .Returns(Result.Success<ServicePrincipalDetails, GraphOperationError>(details));
    if (details.ServicePrincipal.AppId is { } clientId)
      applications.FindByAppIdAsync(clientId, Arg.Any<CancellationToken>())
        .Returns(Result.Success<Maybe<ApplicationDetails>, GraphOperationError>(Maybe<ApplicationDetails>.None));
    var services = new ServiceCollection();
    services.AddLogging();
    services.AddMudServices();
    services.AddSingleton<IJSRuntime>(Substitute.For<IJSRuntime>());
    services.AddSingleton<NavigationManager>(new TestNavigation());
    services.AddSingleton<ICurrentTenantContext>(context.Tenants);
    services.AddSingleton(applications);
    services.AddSingleton(principals);
    services.AddSingleton(policies);
    await using var provider = services.BuildServiceProvider();
    await using var renderer = new HtmlRenderer(provider, provider.GetRequiredService<ILoggerFactory>());
    return await renderer.Dispatcher.InvokeAsync(async () =>
    {
      var parameters = ParameterView.FromDictionary(new Dictionary<string, object?>
      {
        ["ObjectId"] = details.ServicePrincipal.ObjectId.ToString()
      });
      var root = await renderer.RenderComponentAsync<PrincipalPage>(parameters);
      return root.ToHtmlString();
    });
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
