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
using PolicyPage = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin.ClaimsMappingPolicyDetails;
using PoliciesPage = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin.ClaimsMappingPolicies;
using AssignedPolicies = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared.AssignedClaimsMappingPolicies;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingPolicyPageTests
{
  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Policy_list_shows_loading_then_empty_results_or_a_useful_error(bool denied)
  {
    await using var harness = new Harness();
    var pending = new TaskCompletionSource<Result<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>>();
    harness.Service.ListAsync(Arg.Any<CancellationToken>()).Returns(pending.Task);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = harness.Renderer.BeginRenderingComponent<PoliciesPage>();
      Assert.Contains("Loading directory objects", root.ToHtmlString());
      pending.SetResult(denied
        ? Result.Failure<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>(
          new(GraphOperationErrorType.Forbidden, "Check Policy.Read.All admin consent."))
        : Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([]));
      await root.QuiescenceTask;
      var html = root.ToHtmlString();
      Assert.DoesNotContain("Loading directory objects", html);
      Assert.Contains(denied ? "Check Policy.Read.All admin consent" : "No Claims Mapping Policies in the current tenant", html);
    });
  }

  [Fact]
  public async Task Details_render_understood_and_original_definitions_and_clear_every_ID_on_tenant_switch()
  {
    await using var harness = new Harness();
    var policyId = Guid.NewGuid();
    var principalId = Guid.NewGuid();
    var clientId = Guid.NewGuid();
    var applicationId = Guid.NewGuid();
    var policy = new ClaimsMappingPolicyListItem(policyId, "Department policy", false,
      [ClaimsMappingDefinitionParser.Parse(ClaimsMappingDefinitionParserTests.Definition), ClaimsMappingDefinitionParser.Parse("{broken")]);
    var assignments = new ClaimsMappingAssignments([new(principalId, clientId, "Enterprise", null, null, null)],
      [new(applicationId, clientId, "Registration", null, null, null)], []);
    harness.Service.GetByObjectIdAsync(policyId, Arg.Any<CancellationToken>()).Returns(
      Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(new(policy, Result.Success<ClaimsMappingAssignments, GraphOperationError>(assignments))));

    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = await harness.Renderer.RenderComponentAsync<PolicyPage>(Parameters("ObjectId", policyId.ToString()));
      var html = root.ToHtmlString();
      Assert.Contains("Department policy", html);
      Assert.Contains("Version: 1", html);
      Assert.Contains("IncludeBasicClaimSet: True", html);
      Assert.Contains("JWT Claim Type", html);
      Assert.Contains("SAML Claim Type", html);
      Assert.Contains("Transformation Method: Join", html);
      Assert.Contains("Input Parameters", html);
      Assert.Contains("FutureField", html);
      Assert.Contains("FutureRoot", html);
      Assert.Contains("malformed JSON", html);
      Assert.Contains("{broken", html);
      Assert.Contains("IsOrganizationDefault: False", html);
      Assert.Contains("Service Principal Object ID", html);
      Assert.Contains("Application Object ID", html);
      Assert.All(new[] { policyId, principalId, clientId, applicationId }, id => Assert.Contains(id.ToString(), html));
      Assert.DoesNotContain("textarea", html);

      harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
      var cleared = root.ToHtmlString();
      Assert.All(new[] { policyId, principalId, clientId, applicationId }, id => Assert.DoesNotContain(id.ToString(), cleared));
      Assert.DoesNotContain("FutureRoot", cleared);
      Assert.EndsWith("/claims-mapping-policies", harness.Navigation.Uri);
    });
    await harness.Service.Received(1).GetByObjectIdAsync(policyId, Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData("invalid-id")]
  [InlineData("00000000-0000-0000-0000-000000000000")]
  public async Task Invalid_policy_route_is_rejected_without_Graph_discovery(string objectId)
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = await harness.Renderer.RenderComponentAsync<PolicyPage>(Parameters("ObjectId", objectId));
      Assert.Contains("Policy Object ID must be a non-empty GUID", root.ToHtmlString());
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Policy_list_automatically_reloads_the_new_tenant_and_removes_previous_policies()
  {
    await using var harness = new Harness();
    var policyId = Guid.NewGuid();
    harness.Service.ListAsync(Arg.Any<CancellationToken>()).Returns(_ =>
      Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>(
        harness.Context.Tenants.CurrentTenant.Value.TenantId == harness.Context.TenantA
          ? [new(policyId, "Tenant A policy", null, [ClaimsMappingDefinitionParser.Parse(ClaimsMappingDefinitionParserTests.Definition)])] : []));

    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = await harness.Renderer.RenderComponentAsync<PoliciesPage>();
      var html = root.ToHtmlString();
      Assert.Contains("Tenant A policy", html);
      Assert.Contains(policyId.ToString(), html);
      Assert.Contains("Parsed ClaimsSchema Entries", html);
      harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
      html = root.ToHtmlString();
      Assert.Contains("Tenant B", html);
      Assert.DoesNotContain(policyId.ToString(), html);
      Assert.DoesNotContain("Tenant A policy", html);
      Assert.Contains("No Claims Mapping Policies in the current tenant", html);
    });
    await harness.Service.Received(2).ListAsync(Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Assigned_policies_use_the_selected_principal_and_clear_on_tenant_switch()
  {
    await using var harness = new Harness();
    var principalId = Guid.NewGuid();
    var policyId = Guid.NewGuid();
    harness.Service.FindForServicePrincipalAsync(principalId, Arg.Any<CancellationToken>()).Returns(
      Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([new(policyId, "Assigned policy", null, [])]));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = await harness.Renderer.RenderComponentAsync<AssignedPolicies>(Parameters("ServicePrincipalObjectId", principalId));
      Assert.Contains("Assigned policy", root.ToHtmlString());
      Assert.Contains(policyId.ToString(), root.ToHtmlString());
      harness.Context.Tenants.SelectTenant(harness.Context.TenantB);
      Assert.DoesNotContain("Assigned policy", root.ToHtmlString());
      Assert.DoesNotContain(policyId.ToString(), root.ToHtmlString());
    });
    await harness.Service.Received(1).FindForServicePrincipalAsync(principalId, Arg.Any<CancellationToken>());
  }

  private static ParameterView Parameters(string name, object value) =>
    ParameterView.FromDictionary(new Dictionary<string, object?> { [name] = value });

  private sealed class Harness : IAsyncDisposable
  {
    private readonly ServiceProvider _provider;
    public GraphTestContext Context { get; } = new();
    public IClaimsMappingPolicyService Service { get; } = Substitute.For<IClaimsMappingPolicyService>();
    public TestNavigation Navigation { get; } = new();
    public HtmlRenderer Renderer { get; }

    public Harness()
    {
      Context.Tenants.SelectTenant(Context.TenantA);
      var services = new ServiceCollection();
      services.AddLogging();
      services.AddMudServices();
      services.AddSingleton<IJSRuntime>(Substitute.For<IJSRuntime>());
      services.AddSingleton<NavigationManager>(Navigation);
      services.AddSingleton<ICurrentTenantContext>(Context.Tenants);
      services.AddSingleton(Service);
      _provider = services.BuildServiceProvider();
      Renderer = new(_provider, _provider.GetRequiredService<ILoggerFactory>());
    }

    public async ValueTask DisposeAsync()
    {
      await Renderer.DisposeAsync();
      await _provider.DisposeAsync();
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
