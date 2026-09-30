using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Components;

public partial class App
{
    [CascadingParameter]
    public HttpContext? HttpContext { get; set; }

    // The public sign-in page is static; only authenticated users open a circuit.
    private IComponentRenderMode? PageRenderMode => HttpContext?.User.Identity?.IsAuthenticated == true
        ? new InteractiveServerRenderMode(prerender: false)
        : null;
}
