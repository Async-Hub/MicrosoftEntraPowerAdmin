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
using MudBlazor.Extensions;
using MudBlazor.Services;
using NSubstitute;
using PolicyPage = AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components.Pages.Admin.ClaimsMappingPolicyDetails;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class ClaimsMappingPolicyEditingComponentTests
{
  [Theory]
  [InlineData("{broken", "could not safely parse")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"Version\":1,\"IncludeBasicClaimSet\":\"true\",\"ClaimsSchema\":[],\"ClaimsTransformation\":[{\"ID\":\"KeepTransformation\",\"TransformationMethod\":\"Join\"}]}}", "prevent losing")]
  [InlineData("{\"ClaimsMappingPolicy\":{\"FutureProperty\":true}}", "prevent losing")]
  public async Task Unsupported_editor_shows_original_definition_and_name_control_without_definition_controls(string raw, string explanation)
  {
    await using var harness = new Harness(raw);
    await using var renderer = new HtmlRenderer(harness.Provider, harness.Provider.GetRequiredService<ILoggerFactory>());
    await renderer.Dispatcher.InvokeAsync(async () =>
    {
      var root = await renderer.RenderComponentAsync<EditClaimsMappingPolicy>(ParameterView.FromDictionary(
        new Dictionary<string, object?> { ["Context"] = harness.Context }));
      var html = root.ToHtmlString();
      Assert.Contains("Display Name", html);
      Assert.Contains(explanation, html);
      Assert.Contains("Current definition (read-only)", html);
      Assert.DoesNotContain("Add claim", html);
      Assert.DoesNotContain("Include Basic Claim Set", html);
      Assert.DoesNotContain("textarea", html);
      if (raw == "{broken") Assert.Contains(raw, html);
      if (raw.Contains("KeepTransformation"))
      {
        Assert.Contains("KeepTransformation", html);
        Assert.Contains("Transformation Method: Join", html);
      }
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Assigned_definition_change_requires_explicit_confirmation_and_review_has_read_only_preview()
  {
    await using var harness = new Harness(assigned: true);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      Assert.True(harness.Renderer.Button("Save").Disabled);
      await harness.Renderer.Components<MudCheckBox<bool>>().Single(box => box.Label == "Include Basic Claim Set").ValueChanged.InvokeAsync(false);
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Save").Disabled);
      Assert.Contains(harness.Renderer.Components<MudText>(), text => text.Typo == Typo.h6);
      var confirm = harness.Renderer.Components<MudCheckBox<bool>>().Single(box => box.Label?.StartsWith("I confirm") == true);
      await harness.Renderer.Button("Save").OnClick.InvokeAsync();
      Assert.Empty(harness.Service.ReceivedCalls());
      await confirm.ValueChanged.InvokeAsync(true);
      Assert.False(harness.Renderer.Button("Save").Disabled);
      await harness.Renderer.Button("Save").OnClick.InvokeAsync();
      Assert.Same(harness.Authoritative, harness.Saved);
    });
    await harness.Service.Received(1).UpdateAsync(harness.Context,
      Arg.Is<UpdateClaimsMappingPolicyRequest>(request => request.Definition != null && !request.Definition.IncludeBasicClaimSet
        && request.ConfirmedAssignmentIds != null && request.ConfirmedAssignmentIds.SequenceEqual(new[] { harness.PrincipalId })), Arg.Any<CancellationToken>());
  }

  [Theory]
  [InlineData(false)]
  [InlineData(true)]
  public async Task Rename_needs_no_assignment_confirmation_and_never_rewrites_definition(bool supported)
  {
    await using var harness = new Harness(supported ? ClaimsMappingPolicyEditingTests.Supported : "{broken", assigned: true);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      await harness.NameField.ValueChanged.InvokeAsync("Renamed");
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      Assert.DoesNotContain(harness.Renderer.Components<MudCheckBox<bool>>(), box => box.Label?.StartsWith("I confirm") == true);
      Assert.False(harness.Renderer.Button("Save").Disabled);
      await harness.Renderer.Button("Save").OnClick.InvokeAsync();
    });
    await harness.Service.Received(1).UpdateAsync(harness.Context,
      Arg.Is<UpdateClaimsMappingPolicyRequest>(request => request.CreatePatch(harness.Context.Original.Policy).Value.Value.Definition == null
        && request.ConfirmedAssignmentIds == null), Arg.Any<CancellationToken>());
  }

  [Fact]
  public async Task Unchanged_editor_and_cancel_never_call_update_or_mutate_cached_policy()
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Save").Disabled);
      await harness.Renderer.Button("Save").OnClick.InvokeAsync();
      await harness.NameField.ValueChanged.InvokeAsync("Local unsaved name");
      await harness.Renderer.Button("Add claim").OnClick.InvokeAsync();
      await harness.Renderer.Button("Cancel").OnClick.InvokeAsync();
      Assert.True(harness.Cancelled);
      Assert.Equal("Policy", harness.Context.Original.Policy.DisplayName);
      Assert.Equal(ClaimsMappingPolicyEditingTests.Supported, harness.Context.Original.Policy.Definitions[0].Raw);
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Opening_editor_preserves_existing_name_whitespace_and_review_leaves_save_disabled()
  {
    await using var harness = new Harness();
    var context = harness.Context with
    {
      Original = harness.Context.Original with { Policy = harness.Context.Original.Policy with { DisplayName = "  Policy  " } }
    };
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.Renderer.MountAsync<MudPopoverProvider>(ParameterView.Empty);
      await harness.Renderer.MountAsync<EditClaimsMappingPolicy>(ParameterView.FromDictionary(new Dictionary<string, object?> { ["Context"] = context }));
      Assert.Equal("  Policy  ", harness.NameField.GetState(input => input.Value));
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Save").Disabled);
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Shared_claim_controls_add_remove_switch_modes_and_sources_without_changing_original()
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      var editor = harness.Renderer.Components<ClaimsMappingPolicyDefinitionEditor>().Single();
      await harness.Renderer.Components<MudSelect<ClaimValueMode>>().First().ValueChanged.InvokeAsync(ClaimValueMode.ConstantValue);
      Assert.Equal(ClaimValueMode.ConstantValue, editor.Draft.Claims[0].Mode);
      Assert.Equal("", editor.Draft.Claims[0].Id);
      await harness.Renderer.Components<MudTextField<string>>().First(field => field.Label == "Value").ValueChanged.InvokeAsync("West");
      await harness.Renderer.Components<MudSelect<ClaimValueMode>>().First().ValueChanged.InvokeAsync(ClaimValueMode.DirectoryAttribute);
      Assert.Equal("", editor.Draft.Claims[0].Value);
      await harness.Renderer.Components<MudSelect<string>>().First(field => field.Label == "Source").ValueChanged.InvokeAsync("company");
      Assert.Equal("company", editor.Draft.Claims[0].Source);
      await harness.Renderer.Button("Add claim").OnClick.InvokeAsync();
      Assert.Equal(3, editor.Draft.Claims.Count);
      var remove = harness.Renderer.Components<MudButton>().Last(button => harness.Renderer.IsButton(button, "Remove claim"));
      await remove.OnClick.InvokeAsync();
      Assert.Equal(2, editor.Draft.Claims.Count);
      Assert.Equal(ClaimsMappingPolicyEditingTests.Supported, harness.Context.Original.Policy.Definitions[0].Raw);
    });
  }

  [Fact]
  public async Task Busy_save_prevents_duplicate_submissions_and_finishes_once()
  {
    await using var harness = new Harness();
    var pending = new TaskCompletionSource<Result<ClaimsMappingPolicyDetails, GraphOperationError>>();
    harness.Service.UpdateAsync(harness.Context, Arg.Any<UpdateClaimsMappingPolicyRequest>(), Arg.Any<CancellationToken>()).Returns(pending.Task);
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      await harness.NameField.ValueChanged.InvokeAsync("Renamed");
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      var save = harness.Renderer.Button("Save");
      var click = save.OnClick.InvokeAsync();
      Assert.True(save.Disabled);
      await save.OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Cancel").Disabled);
      pending.SetResult(Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(harness.Authoritative));
      await click;
      Assert.Same(harness.Authoritative, harness.Saved);
      await save.OnClick.InvokeAsync();
    });
    await harness.Service.Received(1).UpdateAsync(harness.Context, Arg.Any<UpdateClaimsMappingPolicyRequest>(), Arg.Any<CancellationToken>());
    Assert.Equal(1, harness.SaveNotifications);
  }

  [Theory]
  [InlineData(GraphOperationErrorType.StaleState)]
  [InlineData(GraphOperationErrorType.ConfirmationRequired)]
  public async Task Stale_failures_require_reload_and_discard_review(GraphOperationErrorType type)
  {
    await using var harness = new Harness();
    harness.Service.UpdateAsync(harness.Context, Arg.Any<UpdateClaimsMappingPolicyRequest>(), Arg.Any<CancellationToken>()).Returns(
      Result.Failure<ClaimsMappingPolicyDetails, GraphOperationError>(new(type, "Reload the latest policy")));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      await harness.NameField.ValueChanged.InvokeAsync("Renamed");
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      await harness.Renderer.Button("Save").OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Save").Disabled);
      Assert.True(harness.Renderer.Button("Review changes").Disabled);
      await harness.Renderer.Button("Reload latest (discard edits)").OnClick.InvokeAsync();
      Assert.True(harness.Reloaded);
      Assert.Contains(harness.Renderer.Components<DirectoryStatus>(), status => status.Error?.Type == type);
    });
  }

  [Fact]
  public async Task Deleted_policy_invalidates_editor_and_notifies_parent()
  {
    await using var harness = new Harness();
    harness.Service.UpdateAsync(harness.Context, Arg.Any<UpdateClaimsMappingPolicyRequest>(), Arg.Any<CancellationToken>()).Returns(
      Result.Failure<ClaimsMappingPolicyDetails, GraphOperationError>(new(GraphOperationErrorType.NotFound, "Policy deleted")));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      await harness.NameField.ValueChanged.InvokeAsync("Renamed");
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      await harness.Renderer.Button("Save").OnClick.InvokeAsync();
      Assert.Equal(GraphOperationErrorType.NotFound, harness.Missing?.Type);
      Assert.Empty(harness.Renderer.Components<MudTextField<string>>());
      Assert.Empty(harness.Renderer.Components<ClaimsMappingPolicyDefinitionEditor>());
    });
  }

  [Fact]
  public async Task Tenant_switch_clears_draft_and_blocks_save_even_after_switching_back()
  {
    await using var harness = new Harness();
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.MountAsync();
      await harness.NameField.ValueChanged.InvokeAsync("Renamed");
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      var save = harness.Renderer.Button("Save");
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantB);
      harness.Tenants.Tenants.SelectTenant(harness.Tenants.TenantA);
      Assert.Empty(harness.Renderer.Components<MudTextField<string>>());
      Assert.Empty(harness.Renderer.Components<ClaimsMappingDefinitionView>());
      await save.OnClick.InvokeAsync();
      Assert.Null(harness.Saved);
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Assignment_discovery_failure_allows_rename_but_blocks_definition_save()
  {
    await using var harness = new Harness();
    var context = harness.Context with
    {
      Original = harness.Context.Original with
      {
        Assignments = Result.Failure<ClaimsMappingAssignments, GraphOperationError>(new(GraphOperationErrorType.Forbidden, "Assignments unavailable"))
      }
    };
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.Renderer.MountAsync<MudPopoverProvider>(ParameterView.Empty);
      await harness.Renderer.MountAsync<EditClaimsMappingPolicy>(ParameterView.FromDictionary(new Dictionary<string, object?> { ["Context"] = context }));
      await harness.NameField.ValueChanged.InvokeAsync("Renamed");
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      Assert.False(harness.Renderer.Button("Save").Disabled);
      await harness.Renderer.Components<MudCheckBox<bool>>().Single(box => box.Label == "Include Basic Claim Set").ValueChanged.InvokeAsync(false);
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      Assert.True(harness.Renderer.Button("Save").Disabled);
      Assert.Contains(harness.Renderer.Components<DirectoryStatus>(), status => status.Error?.Message.Contains("Assignments could not be retrieved") == true);
    });
    Assert.Empty(harness.Service.ReceivedCalls());
  }

  [Fact]
  public async Task Details_Edit_loads_fresh_state_then_success_displays_authoritative_update()
  {
    await using var harness = new Harness();
    var fresh = harness.Context with { Original = harness.Context.Original with { Policy = harness.Context.Original.Policy with { DisplayName = "Fresh name" } } };
    harness.Service.GetByObjectIdAsync(harness.Context.Original.Policy.ObjectId, Arg.Any<CancellationToken>()).Returns(
      Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(harness.Context.Original));
    harness.Service.BeginEditAsync(harness.Context.Original.Policy.ObjectId, Arg.Any<CancellationToken>()).Returns(
      Result.Success<ClaimsMappingPolicyEditContext, GraphOperationError>(fresh));
    harness.Service.UpdateAsync(fresh, Arg.Any<UpdateClaimsMappingPolicyRequest>(), Arg.Any<CancellationToken>()).Returns(
      Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(harness.Authoritative));
    await harness.Renderer.Dispatcher.InvokeAsync(async () =>
    {
      await harness.Renderer.MountAsync<MudPopoverProvider>(ParameterView.Empty);
      await harness.Renderer.MountAsync<PolicyPage>(ParameterView.FromDictionary(new Dictionary<string, object?>
      {
        ["ObjectId"] = harness.Context.Original.Policy.ObjectId.ToString()
      }));
      await harness.Renderer.Button("Edit").OnClick.InvokeAsync();
      Assert.Equal("Fresh name", harness.NameField.GetState(input => input.Value));
      await harness.NameField.ValueChanged.InvokeAsync("Draft name");
      await harness.Renderer.Button("Review changes").OnClick.InvokeAsync();
      await harness.Renderer.Button("Save").OnClick.InvokeAsync();
      Assert.Empty(harness.Renderer.Components<EditClaimsMappingPolicy>());
      Assert.Contains(harness.Renderer.Components<MudTable<ServicePrincipalListItem>>().Single().Items ?? [],
        principal => principal.DisplayName == "Graph authoritative assignment");
    });
    await harness.Service.Received(1).BeginEditAsync(harness.Context.Original.Policy.ObjectId, Arg.Any<CancellationToken>());
    await harness.Service.Received(1).GetByObjectIdAsync(harness.Context.Original.Policy.ObjectId, Arg.Any<CancellationToken>());
  }

  private sealed class Harness : IAsyncDisposable
  {
    private readonly ServiceProvider _provider;
    public IServiceProvider Provider => _provider;
    public GraphTestContext Tenants { get; } = new();
    public IClaimsMappingPolicyService Service { get; } = Substitute.For<IClaimsMappingPolicyService>();
    public ClaimsMappingPolicyEditContext Context { get; }
    public ClaimsMappingPolicyDetails Authoritative { get; }
    public Guid PrincipalId { get; } = Guid.NewGuid();
    public InteractionRenderer Renderer { get; }
    public ClaimsMappingPolicyDetails? Saved { get; private set; }
    public GraphOperationError? Missing { get; private set; }
    public bool Cancelled { get; private set; }
    public bool Reloaded { get; private set; }
    public int SaveNotifications { get; private set; }
    public MudTextField<string> NameField => Renderer.Components<MudTextField<string>>().Single(input => input.Label == "Display Name");

    public Harness(string raw = ClaimsMappingPolicyEditingTests.Supported, bool assigned = false)
    {
      Tenants.Tenants.SelectTenant(Tenants.TenantA);
      var principal = new ServicePrincipalListItem(PrincipalId, Guid.NewGuid(), "Assigned Enterprise", null, null, null);
      Context = new(Tenants.TenantA, Guid.NewGuid(), new(ClaimsMappingPolicyEditingTests.Policy(raw),
        Result.Success<ClaimsMappingAssignments, GraphOperationError>(new(assigned ? [principal] : [], [], []))));
      Authoritative = new(Context.Original.Policy with { DisplayName = "Graph authoritative name" },
        Result.Success<ClaimsMappingAssignments, GraphOperationError>(new([principal with { DisplayName = "Graph authoritative assignment" }], [], [])));
      Service.UpdateAsync(Context, Arg.Any<UpdateClaimsMappingPolicyRequest>(), Arg.Any<CancellationToken>()).Returns(
        Result.Success<ClaimsMappingPolicyDetails, GraphOperationError>(Authoritative));
      Service.ClearReceivedCalls();
      var services = new ServiceCollection();
      services.AddLogging();
      services.AddMudServices();
      services.AddSingleton<IJSRuntime>(Substitute.For<IJSRuntime>());
      services.AddSingleton<NavigationManager>(new TestNavigation());
      services.AddSingleton<ICurrentTenantContext>(Tenants.Tenants);
      services.AddSingleton(Service);
      _provider = services.BuildServiceProvider();
      // The renderer's unstable API is confined to this offline component interaction harness.
#pragma warning disable BL0006
      Renderer = new(_provider, _provider.GetRequiredService<ILoggerFactory>());
#pragma warning restore BL0006
    }

    public async Task MountAsync()
    {
      await Renderer.MountAsync<MudPopoverProvider>(ParameterView.Empty);
      await Renderer.MountAsync<EditClaimsMappingPolicy>(ParameterView.FromDictionary(new Dictionary<string, object?>
      {
        ["Context"] = Context,
        ["OnSaved"] = EventCallback.Factory.Create<ClaimsMappingPolicyDetails>(this, details => { Saved = details; SaveNotifications++; }),
        ["OnMissing"] = EventCallback.Factory.Create<GraphOperationError>(this, error => Missing = error),
        ["OnCancel"] = EventCallback.Factory.Create(this, () => Cancelled = true),
        ["OnReload"] = EventCallback.Factory.Create(this, () => Reloaded = true)
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
