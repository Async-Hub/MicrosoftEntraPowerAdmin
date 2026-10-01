namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed record ServicePrincipalListItem(
  Guid ObjectId, Guid? AppId, string DisplayName, string? ServicePrincipalType,
  bool? AccountEnabled, string? PublisherName);

public sealed record ServicePrincipalDetails(
  ServicePrincipalListItem ServicePrincipal, string? Description, string? Homepage,
  string? LoginUrl, IReadOnlyList<string> ServicePrincipalNames,
  Guid? AppOwnerOrganizationId, IReadOnlyList<string> Tags);
