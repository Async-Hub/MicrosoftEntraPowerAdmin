using Microsoft.Extensions.Options;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Tenants;

public sealed class MicrosoftEntraPowerAdminOptionsValidator : IValidateOptions<MicrosoftEntraPowerAdminOptions>
{
    public ValidateOptionsResult Validate(string? name, MicrosoftEntraPowerAdminOptions options)
    {
        if (options.Tenants is not { Length: > 0 })
        {
            return ValidateOptionsResult.Fail("At least one administration tenant must be configured.");
        }

        var errors = new List<string>();
        var ids = new HashSet<Guid>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var tenant in options.Tenants)
        {
            if (tenant is null)
            {
                errors.Add("Administration tenant entries must not be null.");
                continue;
            }

            if (tenant.TenantId == Guid.Empty)
                errors.Add("Administration tenant IDs must not be empty.");
            if (!ids.Add(tenant.TenantId))
                errors.Add("Administration tenant IDs must be unique.");
            if (string.IsNullOrWhiteSpace(tenant.Name))
                errors.Add("Administration tenant names must not be empty.");
            else if (!names.Add(tenant.Name.Trim()))
                errors.Add("Administration tenant names must be unique (ignoring case and surrounding whitespace).");
        }

        return errors.Count == 0 ? ValidateOptionsResult.Success : ValidateOptionsResult.Fail(errors);
    }
}
