namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed record DirectoryPage<T>(IReadOnlyList<T> Items, DirectoryContinuation? NextPage);

// Created only from Graph responses and retained on the server, never accepted from a route.
public sealed class DirectoryContinuation
{
  internal DirectoryContinuation(string url, Guid selectionId)
  {
    Url = url;
    SelectionId = selectionId;
  }

  internal string Url { get; }
  internal Guid SelectionId { get; }
}
