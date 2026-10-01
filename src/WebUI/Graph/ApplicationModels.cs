namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public sealed record ApplicationListItem(
  Guid ObjectId, Guid? AppId, string DisplayName, string? SignInAudience,
  DateTimeOffset? CreatedDateTime, string? PublisherDomain);

public sealed record ApplicationDetails(
  ApplicationListItem Application, string? Description,
  IReadOnlyList<string> IdentifierUris, IReadOnlyList<string> WebRedirectUris,
  IReadOnlyList<string> SpaRedirectUris, IReadOnlyList<string> PublicClientRedirectUris,
  IReadOnlyList<string> Tags, bool? AcceptMappedClaims, int? RequestedAccessTokenVersion);
