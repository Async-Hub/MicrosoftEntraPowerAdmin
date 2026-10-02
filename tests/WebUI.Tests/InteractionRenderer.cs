using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.RenderTree;
using Microsoft.Extensions.Logging;
using MudBlazor;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tests;

// Blazor's renderer API is used only in this offline harness to invoke actual component event callbacks.
#pragma warning disable BL0006
internal sealed class InteractionRenderer(IServiceProvider services, ILoggerFactory logger) : Renderer(services, logger)
{
  private readonly Dictionary<int, IComponent> _components = [];
  public override Dispatcher Dispatcher { get; } = Dispatcher.CreateDefault();
  public Task MountAsync<T>(ParameterView parameters) where T : IComponent
  {
    var component = InstantiateComponent(typeof(T));
    var id = AssignRootComponentId(component);
    _components[id] = component;
    return RenderRootComponentAsync(id, parameters);
  }
  public IEnumerable<T> Components<T>() => _components.Values.OfType<T>();
  public MudButton Button(string label) => Components<MudButton>().Single(button => IsButton(button, label));
  public bool IsButton(MudButton button, string label)
  {
    var id = _components.Single(pair => ReferenceEquals(pair.Value, button)).Key;
    var frames = GetCurrentRenderTreeFrames(id);
    return frames.Array.Take(frames.Count).Any(frame => frame.FrameType == RenderTreeFrameType.Text && frame.TextContent.Trim() == label);
  }
  protected override Task UpdateDisplayAsync(in RenderBatch batch)
  {
    foreach (var frame in batch.ReferenceFrames.Array.Take(batch.ReferenceFrames.Count))
      if (frame.FrameType == RenderTreeFrameType.Component)
        _components[frame.ComponentId] = frame.Component;
    foreach (var id in batch.DisposedComponentIDs.Array.Take(batch.DisposedComponentIDs.Count))
      _components.Remove(id);
    return Task.CompletedTask;
  }
  protected override void HandleException(Exception exception) => System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception).Throw();
}
#pragma warning restore BL0006

