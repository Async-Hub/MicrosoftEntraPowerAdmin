using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Shared;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;
using CSharpFunctionalExtensions;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.JSInterop;
using MudBlazor;
using MudBlazor.Services;
using NSubstitute;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingPolicyAssignmentComponentTests
{
  private static readonly ServicePrincipalListItem Principal = new(Guid.NewGuid(), Guid.NewGuid(), "Enterprise", null, null, null);
  private static readonly ClaimsMappingPolicyListItem PolicyA = new(Guid.NewGuid(), "Policy A", false, []);
  private static readonly ClaimsMappingPolicyListItem PolicyB = new(Guid.NewGuid(), "Policy B", false, []);
  private static readonly ClaimsMappingPolicyListItem PolicyC = new(Guid.NewGuid(), "Policy C", false, []);

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Confirmation_explains_effect_and_identifies_objects_without_mutating(bool unassign)
  {
    await using var harness = new Harness();
    await using var renderer = new HtmlRenderer(harness.Provider, harness.Provider.GetRequiredService<ILoggerFactory>());
    await renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = await renderer.RenderComponentAsync<ChangeClaimsMappingPolicyAssignment>(ParameterView.FromDictionary(
        new Dictionary<string, object?>
        {
          ["Context"] = harness.Context, ["Policy"] = PolicyA,
          ["Principal"] = unassign ? Principal : null, ["IsUnassignment"] = unassign
        }));
      var html = root.ToHtmlString();
      Assert.Contains("Policy A", html);
      Assert.Contains(PolicyA.ObjectId.ToString(), html);
      Assert.Contains("Tenant A", html);
      Assert.Contains(harness.Context.TenantId.ToString(), html);
      if (unassign)
      {
        Assert.Contains(Principal.ObjectId.ToString(), html);
        Assert.Contains("Service Principal Object ID", html);
        Assert.Contains("Unassignment removes only this relationship", html);
        Assert.Contains("does not delete the policy, Service Principal, or App Registration", html);
      }
      else
      {
        Assert.Contains("takes precedence over Custom Claims Policy", html);
        Assert.Contains("Entra admin center", html);
      }
    });
    await harness.Policies.DidNotReceive().AssignAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), Arg.Any<Guid>(), Arg.Any<Guid>(), Arg.Any<CancellationToken>());
    await harness.Policies.DidNotReceive().UnassignAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), Arg.Any<Guid>(), Arg.Any<Guid>(), Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Enterprise_application_view_reloads_Graph_state_after_success_or_stale_failure(bool fails)
  {
    await using var harness = new Harness();
    var changed = false;
    harness.Policies.BeginAssignment().Returns(Result.Success<ClaimsMappingPolicyAssignmentContext, GraphOperationError>(harness.Context));
    harness.Policies.ListAsync(Arg.Any<CancellationToken>()).Returns(Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([PolicyA]));
    harness.Policies.GetAssignedPoliciesAsync(harness.Context, Principal.ObjectId, Arg.Any<CancellationToken>()).Returns(_ =>
      Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>(changed ? [PolicyC] : []));
    harness.Policies.AssignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>()).Returns(_ =>
    {
      changed = true;
      return fails ? UnitResult.Failure(new GraphOperationError(GraphOperationErrorType.InvalidInput, "Stale relationship"))
        : UnitResult.Success<GraphOperationError>();
    });
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.Renderer.MountAsync<MudPopoverProvider>(ParameterView.Empty);
      await harness.Renderer.MountAsync<AssignedClaimsMappingPolicies>(ParameterView.FromDictionary(new Dictionary<string, object?> { ["ServicePrincipal"] = Principal }));
      await harness.Renderer.Button("Assign policy").OnClick.InvokeAsync();
      await harness.Renderer.Components<MudSelect<ClaimsMappingPolicyListItem>>().Single().ValueChanged.InvokeAsync(PolicyA);
      await harness.Renderer.Button("Assign").OnClick.InvokeAsync();
      Assert.Equal(new[] { PolicyC }, harness.Renderer.Components<MudTable<ClaimsMappingPolicyListItem>>().Single().Items);
      Assert.Empty(harness.Renderer.Components<ChangeClaimsMappingPolicyAssignment>());
      if (fails)
        Assert.Contains(harness.Renderer.Components<DirectoryStatus>(), status => status.Error?.Message == "Stale relationship");
    });
    await harness.Policies.Received(3).GetAssignedPoliciesAsync(harness.Context, Principal.ObjectId, Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Policy_details_reload_appliesTo_after_assignment_and_unassignment(bool unassign)
  {
    await using var harness = new Harness();
    var changed = false;
    var graphPrincipal = Principal with { ObjectId = Guid.NewGuid(), DisplayName = "Graph response" };
    harness.Policies.BeginAssignment().Returns(Result.Success<ClaimsMappingPolicyAssignmentContext, GraphOperationError>(harness.Context));
    harness.Policies.GetByObjectIdAsync(PolicyA.ObjectId, Arg.Any<CancellationToken>()).Returns(_ =>
      Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(new(PolicyA,
        Result.Success<ClaimsMappingAssignments, GraphOperationError>(new(changed ? [graphPrincipal] : unassign ? [Principal] : [], [], [])))));
    harness.Principals.SearchAsync(Arg.Any<string>(), null, Arg.Any<CancellationToken>()).Returns(
      Result.Success<DirectoryPage<ServicePrincipalListItem>, GraphOperationError>(new([Principal], null)));
    harness.Policies.AssignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>()).Returns(_ =>
    {
      changed = true;
      return UnitResult.Success<GraphOperationError>();
    });
    harness.Policies.UnassignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>()).Returns(_ =>
    {
      changed = true;
      return UnitResult.Success<GraphOperationError>();
    });
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.Renderer.MountAsync<MudPopoverProvider>(ParameterView.Empty);
      await harness.Renderer.MountAsync<Components.Pages.Admin.ClaimsMappingPolicyDetails>(ParameterView.FromDictionary(
        new Dictionary<string, object?> { ["ObjectId"] = PolicyA.ObjectId.ToString() }));
      await harness.Renderer.Button(unassign ? "Unassign" : "Assign to Enterprise Application").OnClick.InvokeAsync();
      if (!unassign)
        await harness.Renderer.Button("Select").OnClick.InvokeAsync();
      // The policy-side Unassign button remains present but disabled behind the confirmation.
      var submit = harness.Renderer.Components<MudButton>().Single(button => harness.Renderer.IsButton(button, unassign ? "Unassign" : "Assign") && !button.Disabled);
      await submit.OnClick.InvokeAsync();
      Assert.Equal(new[] { graphPrincipal }, harness.Renderer.Components<MudTable<ServicePrincipalListItem>>().Single().Items);
      Assert.Empty(harness.Renderer.Components<ChangeClaimsMappingPolicyAssignment>());
    });
    await harness.Policies.Received(2).GetByObjectIdAsync(PolicyA.ObjectId, Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Principal_picker_excludes_assigned_policies_and_selection_requires_explicit_confirmation()
  {
    await using var harness = new Harness();
    harness.Policies.ListAsync(Arg.Any<CancellationToken>()).Returns(
      Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([PolicyA, PolicyB, PolicyC]));
    harness.Policies.GetAssignedPoliciesAsync(harness.Context, Principal.ObjectId, Arg.Any<CancellationToken>()).Returns(
      Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([PolicyB]));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync(Principal, null);
      Assert.Equal(new[] { PolicyA, PolicyC }, harness.Renderer.Components<MudSelectItem<ClaimsMappingPolicyListItem>>().Select(item => item.Value));
      Assert.True(harness.Renderer.Button("Assign").Disabled);
      await harness.Renderer.Components<MudSelect<ClaimsMappingPolicyListItem>>().Single().ValueChanged.InvokeAsync(PolicyC);
      Assert.False(harness.Renderer.Button("Assign").Disabled);
      await harness.Policies.DidNotReceive().AssignAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), Arg.Any<Guid>(), Arg.Any<Guid>(), Arg.Any<CancellationToken>());
      await harness.Renderer.Button("Assign").OnClick.InvokeAsync();
      await harness.Policies.Received(1).AssignAsync(harness.Context, Principal.ObjectId, PolicyC.ObjectId, Arg.Any<CancellationToken>());
      Assert.True(harness.Finished);
    });
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Busy_confirmation_prevents_duplicate_mutations_and_finishes_once(bool unassign)
  {
    await using var harness = new Harness();
    var pending = new TaskCompletionSource<UnitResult<GraphOperationError>>();
    if (unassign)
      harness.Policies.UnassignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>()).Returns(pending.Task);
    else
    {
      harness.Policies.ListAsync(Arg.Any<CancellationToken>()).Returns(Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([PolicyA]));
      harness.Policies.AssignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>()).Returns(pending.Task);
    }
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync(Principal, unassign ? PolicyA : null, unassign);
      if (!unassign)
        await harness.Renderer.Components<MudSelect<ClaimsMappingPolicyListItem>>().Single().ValueChanged.InvokeAsync(PolicyA);
      var button = harness.Renderer.Button(unassign ? "Unassign" : "Assign");
      var click = button.OnClick.InvokeAsync();
      Assert.True(button.Disabled);
      await button.OnClick.InvokeAsync();
      Assert.False(harness.Finished);
      pending.SetResult(UnitResult.Success<GraphOperationError>());
      await click;
      Assert.True(harness.Finished);
    });
    if (unassign)
      await harness.Policies.Received(1).UnassignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>());
    else
      await harness.Policies.Received(1).AssignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Tenant_switch_invalidates_confirmation_and_discards_selections_even_after_switching_back(bool unassign)
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync(unassign ? Principal : null, PolicyA, unassign);
      var button = harness.Renderer.Button(unassign ? "Unassign" : "Assign");
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantA);
      Assert.True(button.Disabled);
      Assert.Empty(harness.Renderer.Components<MudSelect<ClaimsMappingPolicyListItem>>());
      await button.OnClick.InvokeAsync();
      Assert.False(harness.Finished);
    });
    await harness.Policies.DidNotReceive().AssignAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), Arg.Any<Guid>(), Arg.Any<Guid>(), Arg.Any<CancellationToken>());
    await harness.Policies.DidNotReceive().UnassignAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), Arg.Any<Guid>(), Arg.Any<Guid>(), Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Policy_picker_reuses_principal_search_excludes_assignments_and_uses_selected_object_ID()
  {
    await using var harness = new Harness();
    var assigned = Principal with { ObjectId = Guid.NewGuid(), DisplayName = "Already assigned" };
    harness.Principals.SearchAsync(Arg.Any<string>(), null, Arg.Any<CancellationToken>()).Returns(
      Result.Success<DirectoryPage<ServicePrincipalListItem>, GraphOperationError>(new([assigned, Principal], null)));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync(null, PolicyA, assignedIds: [assigned.ObjectId]);
      Assert.Single(harness.Renderer.Components<MudButton>(), button => harness.Renderer.IsButton(button, "Select"));
      await harness.Renderer.Button("Select").OnClick.InvokeAsync();
      await harness.Policies.DidNotReceive().AssignAsync(Arg.Any<ClaimsMappingPolicyAssignmentContext>(), Arg.Any<Guid>(), Arg.Any<Guid>(), Arg.Any<CancellationToken>());
      await harness.Renderer.Button("Assign").OnClick.InvokeAsync();
    });
    await harness.Policies.Received(1).AssignAsync(harness.Context, Principal.ObjectId, PolicyA.ObjectId, Arg.Any<CancellationToken>());
  }

  private sealed class Harness : IAsyncDisposable
  {
    private readonly ServiceProvider _provider;
    public IServiceProvider Provider => _provider;
    public GraphTestContext Tenants { get; } = new();
    public ClaimsMappingPolicyAssignmentContext Context { get; }
    public IClaimsMappingPolicyService Policies { get; } = Substitute.For<IClaimsMappingPolicyService>();
    public IServicePrincipalService Principals { get; } = Substitute.For<IServicePrincipalService>();
    public InteractionRenderer Renderer { get; }
    public bool Finished { get; private set; }

    public Harness()
    {
      Tenants.Tenants.SelectTenant(Tenants.TenantA);
      Context = new(Tenants.TenantA, Guid.NewGuid());
      Policies.GetAssignedPoliciesAsync(Context, Arg.Any<Guid>(), Arg.Any<CancellationToken>()).Returns(
        Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([]));
      Principals.SearchAsync(Arg.Any<string>(), null, Arg.Any<CancellationToken>()).Returns(
        Result.Success<DirectoryPage<ServicePrincipalListItem>, GraphOperationError>(new([], null)));
      var services = new ServiceCollection();
      services.AddLogging();
      services.AddMudServices();
      services.AddSingleton<IJSRuntime>(Substitute.For<IJSRuntime>());
      services.AddSingleton<NavigationManager>(new TestNavigation());
      services.AddSingleton<ICurrentTenantContext>(Tenants.Tenants);
      services.AddSingleton(Policies);
      services.AddSingleton(Principals);
      _provider = services.BuildServiceProvider();
      // The renderer's unstable API is confined to this offline component interaction harness.
#pragma warning disable BL0006
      Renderer = new(_provider, _provider.GetRequiredService<ILoggerFactory>());
#pragma warning restore BL0006
    }
    public async Task MountAsync(ServicePrincipalListItem? principal, ClaimsMappingPolicyListItem? policy,
      bool unassign = false, IReadOnlyList<Guid>? assignedIds = null)
    {
      await Renderer.MountAsync<MudPopoverProvider>(ParameterView.Empty);
      await Renderer.MountAsync<ChangeClaimsMappingPolicyAssignment>(ParameterView.FromDictionary(new Dictionary<string, object?>
      {
        ["Context"] = Context, ["Principal"] = principal, ["Policy"] = policy,
        ["IsUnassignment"] = unassign, ["AssignedPrincipalIds"] = assignedIds ?? [],
        ["OnFinished"] = EventCallback.Factory.Create<GraphOperationError?>(this, _ => Finished = true)
      }));
    }
    public async ValueTask DisposeAsync()
    {
#pragma warning disable BL0006
      await Renderer.DisposeAsync();
#pragma warning restore BL0006
      await _provider.DisposeAsync();
    }
  }

  private sealed class TestNavigation : NavigationManager
  {
    public TestNavigation() => Initialize("https://localhost/", "https://localhost/");
    protected override void NavigateToCore(string uri, bool forceLoad) => Uri = ToAbsoluteUri(uri).ToString();
  }
}
