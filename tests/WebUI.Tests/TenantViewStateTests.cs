using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components;
using AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;
using CSharpFunctionalExtensions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

public sealed class TenantViewStateTests
{
  [Fact]
  public async Task Tenant_switch_clears_loaded_model_before_notifying_page()
  {
    var context = new GraphTestContext();
    context.Tenants.SelectTenant(context.TenantA);
    var notified = false;
    using var state = new TenantViewState<string>(context.Tenants, () => notified = true);
    Assert.False(state.HasLoaded);
    await state.LoadAsync(_ => Task.FromResult(Result.Success<string, GraphOperationError>("Tenant A object")));
    Assert.Equal("Tenant A object", state.Value);
    context.Tenants.SelectTenant(context.TenantB);
    Assert.Null(state.Value);
    Assert.Null(state.Error);
    Assert.False(state.HasLoaded);
    Assert.False(state.IsLoading);
    Assert.True(notified);
  }

  [Fact]
  public async Task A_B_A_switch_cancels_pending_load_and_ignores_late_result()
  {
    var context = new GraphTestContext();
    context.Tenants.SelectTenant(context.TenantA);
    using var state = new TenantViewState<string>(context.Tenants, () => { });
    var pending = new TaskCompletionSource<Result<string, GraphOperationError>>();
    CancellationToken received = default;
    var load = state.LoadAsync(token => { received = token; return pending.Task; });
    Assert.True(state.IsLoading);
    context.Tenants.SelectTenant(context.TenantB);
    context.Tenants.SelectTenant(context.TenantA);
    Assert.True(received.IsCancellationRequested);
    pending.SetResult(Result.Success<string, GraphOperationError>("Stale A"));
    await load;
    Assert.Null(state.Value);
    Assert.False(state.HasLoaded);
  }

  [Fact]
  public async Task New_search_supersedes_pending_search_without_overwriting_new_results()
  {
    var context = new GraphTestContext();
    using var state = new TenantViewState<string>(context.Tenants, () => { });
    var pending = new TaskCompletionSource<Result<string, GraphOperationError>>();
    var old = state.LoadAsync(_ => pending.Task);
    await state.LoadAsync(_ => Task.FromResult(Result.Success<string, GraphOperationError>("New results")));
    pending.SetResult(Result.Success<string, GraphOperationError>("Old results"));
    await old;
    Assert.Equal("New results", state.Value);
    Assert.True(state.HasLoaded);
    Assert.False(state.IsLoading);
  }

  [Fact]
  public async Task Failure_clears_previous_data_and_disposal_unsubscribes()
  {
    var context = new GraphTestContext();
    var notifications = 0;
    var state = new TenantViewState<string>(context.Tenants, () => notifications++);
    await state.LoadAsync(_ => Task.FromResult(Result.Success<string, GraphOperationError>("Loaded")));
    var error = new GraphOperationError(GraphOperationErrorType.Forbidden, "Access denied");
    await state.LoadAsync(_ => Task.FromResult(Result.Failure<string, GraphOperationError>(error)));
    Assert.Null(state.Value);
    Assert.Equal(error, state.Error);
    Assert.True(state.HasLoaded);
    state.Dispose();
    context.Tenants.SelectTenant(context.TenantB);
    Assert.Equal(0, notifications);
  }
}
