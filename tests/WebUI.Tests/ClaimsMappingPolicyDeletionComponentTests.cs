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
using PolicyPage = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin.ClaimsMappingPolicyDetails;
using PolicyListPage = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin.ClaimsMappingPolicies;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingPolicyDeletionComponentTests
{
  [Theory]
  [InlineData(ClaimsMappingPolicyEditingTests.Supported)]
  [InlineData("{broken")]
  public async Task Confirmation_identifies_policy_tenant_irreversibility_and_available_definition_summary(string raw)
  {
    await using var harness = new Harness(raw);
    await using var renderer = new HtmlRenderer(harness.Provider, harness.Provider.GetRequiredService<ILoggerFactory>());
    await renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = await renderer.RenderComponentAsync<DeleteClaimsMappingPolicyDialog>(ParameterView.FromDictionary(
        new Dictionary<string, object?> { ["Context"] = harness.Context }));
      var html = root.ToHtmlString();
      Assert.Contains("role=\"dialog\"", html);
      Assert.Contains("Delete Claims Mapping Policy?", html);
      Assert.Contains(harness.Context.Original.Policy.DisplayName, html);
      Assert.Contains(harness.Context.Original.Policy.ObjectId.ToString(), html);
      Assert.Contains("Tenant A", html);
      Assert.Contains(harness.Context.TenantId.ToString(), html);
      Assert.Contains("permanently deletes", html);
      Assert.Contains("cannot be undone", html);
      Assert.Contains("Assignments: 0", html);
      Assert.Contains("Claims Schema entries:", html);
      Assert.Contains("Claims Transformations:", html);
      Assert.Contains("Include Basic Claim Set:", html);
      Assert.DoesNotContain(raw, html);
      if (raw == "{broken") Assert.Contains("Unavailable", html);
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Exact_name_confirmation_enables_final_delete_and_invalid_input_is_rejected()
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountDialogAsync();
      Assert.True(harness.Renderer.Button("Delete policy").Disabled);
      foreach (var invalid in new[] { "Other", "policy", "" })
      {
        await harness.ConfirmationField.ValueChanged.InvokeAsync(invalid);
        var delete = harness.Renderer.Button("Delete policy");
        Assert.True(delete.Disabled);
        await delete.OnClick.InvokeAsync();
      }
      Assert.Empty(harness.Service.ReceivedCalls());
      await harness.ConfirmationField.ValueChanged.InvokeAsync("  Policy  ");
      Assert.False(harness.Renderer.Button("Delete policy").Disabled);
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.Equal(1, harness.Notifications);
      Assert.Null(harness.FinishedError);
    });
    await harness.Service.Received(1).DeleteAsync(harness.Context, "  Policy  ", Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Missing_name_requires_explicit_DELETE_fallback()
  {
    await using var harness = new Harness();
    harness.Context = harness.Context with { CurrentDisplayName = null };
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountDialogAsync();
      Assert.True(harness.Renderer.Button("Delete policy").Disabled);
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      Assert.True(harness.Renderer.Button("Delete policy").Disabled);
      await harness.ConfirmationField.ValueChanged.InvokeAsync("DELETE");
      Assert.False(harness.Renderer.Button("Delete policy").Disabled);
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
    });
    await harness.Service.Received(1).DeleteAsync(harness.Context, "DELETE", Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData(1)]
  [InlineData(3)]
  public async Task Assigned_context_never_enables_final_delete(int assignmentCount)
  {
    await using var harness = new Harness(assignmentCount: assignmentCount);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountDialogAsync();
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      var delete = harness.Renderer.Button("Delete policy");
      Assert.True(delete.Disabled);
      await delete.OnClick.InvokeAsync();
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Busy_delete_disables_cancel_and_prevents_duplicate_submissions_or_notifications()
  {
    await using var harness = new Harness();
    var pending = new TaskCompletionSource<UnitResult<GraphOperationError>>();
    harness.Service.DeleteAsync(harness.Context, Arg.Any<string>(), Arg.Any<CancellationToken>()).Returns(pending.Task);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountDialogAsync();
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      var delete = harness.Renderer.Button("Delete policy");
      var attempt = delete.OnClick.InvokeAsync();
      Assert.True(delete.Disabled);
      Assert.True(harness.ConfirmationField.Disabled);
      Assert.True(harness.Renderer.Button("Cancel").Disabled);
      await harness.Renderer.Button("Cancel").OnClick.InvokeAsync();
      Assert.False(harness.Cancelled);
      await delete.OnClick.InvokeAsync();
      pending.SetResult(UnitResult.Success<GraphOperationError>());
      await attempt;
      await delete.OnClick.InvokeAsync();
      Assert.Equal(1, harness.Notifications);
    });
    await harness.Service.Received(1).DeleteAsync(harness.Context, "Policy", Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Tenant_switch_discards_text_and_invalidates_confirmation_even_after_switching_back()
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountDialogAsync();
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      var delete = harness.Renderer.Button("Delete policy");
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantA);
      Assert.True(delete.Disabled);
      Assert.Empty(harness.Renderer.Components<MudTextField<string>>());
      await delete.OnClick.InvokeAsync();
      Assert.Equal(0, harness.Notifications);
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Tenant_switch_cancels_inflight_delete_and_ignores_late_success()
  {
    await using var harness = new Harness();
    var pending = new TaskCompletionSource<UnitResult<GraphOperationError>>();
    CancellationToken submitted = default;
    harness.Service.DeleteAsync(harness.Context, Arg.Any<string>(), Arg.Any<CancellationToken>()).Returns(call =>
    {
      submitted = call.Arg<CancellationToken>();
      return pending.Task;
    });
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountDialogAsync();
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      var attempt = harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      Assert.True(submitted.IsCancellationRequested);
      pending.SetResult(UnitResult.Success<GraphOperationError>());
      await attempt;
      Assert.Equal(0, harness.Notifications);
    });
  }

  [Fact]
  public async Task Details_fetch_fresh_identity_before_confirmation_and_block_other_mutations_during_delete()
  {
    await using var harness = new Harness();
    var fresh = harness.Context with
    {
      CurrentDisplayName = "Fresh name",
      Original = harness.Context.Original with { Policy = harness.Context.Original.Policy with { DisplayName = "Fresh name" } }
    };
    harness.Service.BeginDeleteAsync(harness.Id, Arg.Any<CancellationToken>()).Returns(
      Result.Success<ClaimsMappingPolicyDeletionContext, GraphOperationError>(fresh));
    var pending = new TaskCompletionSource<UnitResult<GraphOperationError>>();
    harness.Service.DeleteAsync(fresh, Arg.Any<string>(), Arg.Any<CancellationToken>()).Returns(pending.Task);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountPageAsync();
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.Same(fresh, harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>().Single().Context);
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      Assert.True(harness.Renderer.Button("Delete policy").Disabled);
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Fresh name");
      var attempt = harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Edit").Disabled);
      Assert.True(harness.Renderer.Button("Assign to Enterprise Application").Disabled);
      Assert.True(harness.Renderer.Button("Refresh").Disabled);
      await harness.Renderer.Button("Edit").OnClick.InvokeAsync();
      await harness.Renderer.Button("Assign to Enterprise Application").OnClick.InvokeAsync();
      pending.SetResult(UnitResult.Success<GraphOperationError>());
      await attempt;
      Assert.EndsWith("/claims-mapping-policies", harness.Navigation.Uri);
    });
    await harness.Service.Received(1).BeginDeleteAsync(harness.Id, Arg.Any<CancellationToken>());
    await harness.Service.DidNotReceive().BeginEditAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>());
    harness.Service.DidNotReceive().BeginAssignment();
  }

  [Fact]
  public async Task Assigned_details_reuse_Step_7_unassignment_then_allow_deletion_after_final_relationship_removed()
  {
    await using var harness = new Harness(assignmentCount: 1);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountPageAsync();
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.Empty(harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>());
      var principal = Assert.Single(harness.Renderer.Components<MudTable<ServicePrincipalListItem>>().Single().Items ?? []);
      Assert.Equal(harness.Principal.ObjectId, principal.ObjectId);
      var unassigned = harness.Context with
      {
        Original = harness.Context.Original with
        {
          Assignments = Result.Success<ClaimsMappingAssignments, GraphOperationError>(new([], [], []))
        }
      };
      harness.Service.GetByObjectIdAsync(harness.Id, Arg.Any<CancellationToken>()).Returns(
        Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(unassigned.Original));
      harness.Service.BeginDeleteAsync(harness.Id, Arg.Any<CancellationToken>()).Returns(
        Result.Success<ClaimsMappingPolicyDeletionContext, GraphOperationError>(unassigned));
      await harness.Renderer.Button("Unassign").OnClick.InvokeAsync();
      var confirmation = Assert.Single(harness.Renderer.Components<ChangeClaimsMappingPolicyAssignment>());
      Assert.True(confirmation.IsUnassignment);
      Assert.Equal(harness.Principal, confirmation.Principal);
      await harness.Renderer.Components<MudButton>().Last(button => harness.Renderer.IsButton(button, "Unassign")).OnClick.InvokeAsync();
      Assert.Empty(harness.Renderer.Components<ChangeClaimsMappingPolicyAssignment>());
      Assert.Empty(harness.Renderer.Components<MudTable<ServicePrincipalListItem>>().Single().Items ?? []);
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.Equal(0, Assert.Single(harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>()).Context.AssignmentCount);
    });
    await harness.Service.Received(1).UnassignAsync(harness.AssignmentContext, harness.Principal.ObjectId, harness.Id, Arg.Any<CancellationToken>());
    await harness.Service.DidNotReceive().DeleteAsync(Arg.Any<ClaimsMappingPolicyDeletionContext>(), Arg.Any<string>(), Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData(GraphOperationErrorType.PolicyIsAssigned)]
  [InlineData(GraphOperationErrorType.GraphFailure)]
  [InlineData(GraphOperationErrorType.Throttled)]
  [InlineData(GraphOperationErrorType.Forbidden)]
  public async Task Failed_delete_closes_confirmation_refreshes_details_and_preserves_resource(GraphOperationErrorType type)
  {
    await using var harness = new Harness();
    harness.Service.DeleteAsync(harness.Context, Arg.Any<string>(), Arg.Any<CancellationToken>()).Returns(
      UnitResult.Failure(new GraphOperationError(type, "Deletion blocked or unconfirmed")));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountPageAsync();
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      var reloaded = harness.Context.Original with
      {
        Assignments = Result.Success<ClaimsMappingAssignments, GraphOperationError>(new(type == GraphOperationErrorType.PolicyIsAssigned ? [harness.Principal] : [], [], []))
      };
      harness.Service.GetByObjectIdAsync(harness.Id, Arg.Any<CancellationToken>()).Returns(
        Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(reloaded));
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.Empty(harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>());
      Assert.Contains(harness.Renderer.Components<DirectoryStatus>(), status => status.Error?.Type == type);
      Assert.Equal(type == GraphOperationErrorType.PolicyIsAssigned ? 1 : 0,
        harness.Renderer.Components<MudTable<ServicePrincipalListItem>>().Single().Items?.Count() ?? 0);
      Assert.EndsWith(harness.Id.ToString(), harness.Navigation.Uri);
    });
    await harness.Service.Received(2).GetByObjectIdAsync(harness.Id, Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Success_or_DELETE_404_clears_details_navigates_and_list_initialization_reloads_Graph(bool missing)
  {
    await using var harness = new Harness();
    if (missing) harness.Service.DeleteAsync(harness.Context, Arg.Any<string>(), Arg.Any<CancellationToken>()).Returns(
      UnitResult.Failure(new GraphOperationError(GraphOperationErrorType.NotFound, "This policy no longer exists.")));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountPageAsync();
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.Empty(harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>());
      Assert.Empty(harness.Renderer.Components<MudTable<ServicePrincipalListItem>>());
      Assert.EndsWith("/claims-mapping-policies", harness.Navigation.Uri);
      harness.Snackbar.Received(1).Add(missing ? "This policy no longer exists." : "Claims Mapping Policy deleted.",
        missing ? Severity.Info : Severity.Success, Arg.Any<Action<SnackbarOptions>?>(), Arg.Any<string?>());
      await harness.Renderer.MountAsync<PolicyListPage>(ParameterView.Empty);
      Assert.Empty(harness.Renderer.Components<ClaimsMappingPolicyTable>().Single().Policies);
    });
    await harness.Service.Received(1).GetByObjectIdAsync(harness.Id, Arg.Any<CancellationToken>());
    await harness.Service.Received(1).ListAsync(Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Missing_policy_before_confirmation_navigates_to_list_without_delete()
  {
    await using var harness = new Harness();
    harness.Service.BeginDeleteAsync(harness.Id, Arg.Any<CancellationToken>()).Returns(
      Result.Failure<ClaimsMappingPolicyDeletionContext, GraphOperationError>(new(GraphOperationErrorType.NotFound, "This policy no longer exists.")));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountPageAsync();
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.Empty(harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>());
      Assert.EndsWith("/claims-mapping-policies", harness.Navigation.Uri);
    });
    await harness.Service.DidNotReceive().DeleteAsync(Arg.Any<ClaimsMappingPolicyDeletionContext>(), Arg.Any<string>(), Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Tenant_switch_closes_details_confirmation_without_using_disposed_cancellation_state()
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountPageAsync();
      await harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      await harness.ConfirmationField.ValueChanged.InvokeAsync("Policy");
      var delete = harness.Renderer.Button("Delete policy");
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      Assert.Empty(harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>());
      Assert.EndsWith("/claims-mapping-policies", harness.Navigation.Uri);
      await delete.OnClick.InvokeAsync();
    });
    await harness.Service.DidNotReceive().DeleteAsync(Arg.Any<ClaimsMappingPolicyDeletionContext>(), Arg.Any<string>(), Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Tenant_switch_during_confirmation_load_discards_result_and_cancels_request()
  {
    await using var harness = new Harness();
    var pending = new TaskCompletionSource<Result<ClaimsMappingPolicyDeletionContext, GraphOperationError>>();
    CancellationToken submitted = default;
    harness.Service.BeginDeleteAsync(harness.Id, Arg.Any<CancellationToken>()).Returns(call =>
    {
      submitted = call.Arg<CancellationToken>();
      return pending.Task;
    });
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountPageAsync();
      var attempt = harness.Renderer.Button("Delete policy").OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Edit").Disabled);
      Assert.True(harness.Renderer.Button("Delete policy").Disabled);
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      Assert.True(submitted.IsCancellationRequested);
      pending.SetResult(Result.Success<ClaimsMappingPolicyDeletionContext, GraphOperationError>(harness.Context));
      await attempt;
      Assert.Empty(harness.Renderer.Components<DeleteClaimsMappingPolicyDialog>());
      Assert.EndsWith("/claims-mapping-policies", harness.Navigation.Uri);
    });
  }

  private sealed class Harness : IAsyncDisposable
  {
    private readonly ServiceProvider _provider;
    public IServiceProvider Provider => _provider;
    public GraphTestContext Tenants { get; } = new();
    public IClaimsMappingPolicyService Service { get; } = Substitute.For<IClaimsMappingPolicyService>();
    public ISnackbar Snackbar { get; } = Substitute.For<ISnackbar>();
    public TestNavigation Navigation { get; } = new();
    public ClaimsMappingPolicyDeletionContext Context { get; set; }
    public ClaimsMappingPolicyAssignmentContext AssignmentContext { get; }
    public ServicePrincipalListItem Principal { get; } = new(Guid.NewGuid(), Guid.NewGuid(), "Assigned Enterprise", null, null, null);
    public Guid Id => Context.Original.Policy.ObjectId;
    public InteractionRenderer Renderer { get; }
    public GraphOperationError? FinishedError { get; private set; }
    public int Notifications { get; private set; }
    public bool Cancelled { get; private set; }
    public MudTextField<string> ConfirmationField => Renderer.Components<MudTextField<string>>().Single(input => input.Label == "Policy name confirmation");

    public Harness(string raw = ClaimsMappingPolicyEditingTests.Supported, int assignmentCount = 0)
    {
      Tenants.Tenants.SelectTenant(Tenants.TenantA);
      Context = new(Tenants.TenantA, Guid.NewGuid(), new(ClaimsMappingPolicyEditingTests.Policy(raw),
        Result.Success<ClaimsMappingAssignments, GraphOperationError>(new(Enumerable.Range(0, assignmentCount)
          .Select(index => Principal with { ObjectId = index == 0 ? Principal.ObjectId : Guid.NewGuid() }).ToArray(), [], []))), "Policy");
      AssignmentContext = new(Context.TenantId, Context.SelectionId);
      Service.BeginAssignment().Returns(Result.Success<ClaimsMappingPolicyAssignmentContext, GraphOperationError>(AssignmentContext));
      Service.UnassignAsync(AssignmentContext, Principal.ObjectId, Id, Arg.Any<CancellationToken>()).Returns(UnitResult.Success<GraphOperationError>());
      Service.GetByObjectIdAsync(Id, Arg.Any<CancellationToken>()).Returns(Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(Context.Original));
      Service.BeginDeleteAsync(Id, Arg.Any<CancellationToken>()).Returns(Result.Success<ClaimsMappingPolicyDeletionContext, GraphOperationError>(Context));
      Service.DeleteAsync(Context, Arg.Any<string>(), Arg.Any<CancellationToken>()).Returns(UnitResult.Success<GraphOperationError>());
      Service.ListAsync(Arg.Any<CancellationToken>()).Returns(Result.Success<IReadOnlyList<ClaimsMappingPolicyListItem>, GraphOperationError>([]));
      Service.ClearReceivedCalls();
      var services = new ServiceCollection();
      services.AddLogging();
      services.AddMudServices();
      services.AddSingleton(Snackbar);
      services.AddSingleton<IJSRuntime>(Substitute.For<IJSRuntime>());
      services.AddSingleton<NavigationManager>(Navigation);
      services.AddSingleton<ICurrentTenantContext>(Tenants.Tenants);
      services.AddSingleton(Service);
      services.AddSingleton(Substitute.For<IServicePrincipalService>());
      _provider = services.BuildServiceProvider();
#pragma warning disable BL0006
      Renderer = new(_provider, _provider.GetRequiredService<ILoggerFactory>());
#pragma warning restore BL0006
    }

    public Task MountDialogAsync() => Renderer.MountAsync<DeleteClaimsMappingPolicyDialog>(ParameterView.FromDictionary(new Dictionary<string, object?>
    {
      ["Context"] = Context,
      ["OnFinished"] = EventCallback.Factory.Create<GraphOperationError?>(this, error => { FinishedError = error; Notifications++; }),
      ["OnCancel"] = EventCallback.Factory.Create(this, () => Cancelled = true)
    }));

    public Task MountPageAsync()
    {
      Navigation.SetUri($"claims-mapping-policies/{Id:D}");
      return Renderer.MountAsync<PolicyPage>(ParameterView.FromDictionary(new Dictionary<string, object?> { ["ObjectId"] = Id.ToString() }));
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
    public void SetUri(string uri) => Uri = ToAbsoluteUri(uri).ToString();
    protected override void NavigateToCore(string uri, bool forceLoad) => SetUri(uri);
  }
}
