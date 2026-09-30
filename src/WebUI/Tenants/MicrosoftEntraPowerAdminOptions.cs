namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public sealed class MicrosoftEntraPowerAdminOptions
{
    public const string SectionName = "MicrosoftEntraPowerAdmin";

    public MicrosoftEntraTenantOptions[] Tenants { get; init; } = [];
}

public sealed class MicrosoftEntraTenantOptions
{
    public string Name { get; init; } = string.Empty;

    public Guid TenantId { get; init; }
}
