using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions;

namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

internal static class GraphErrorMapping
{
  public static GraphOperationError From(MsalUiRequiredException exception) => new(
    exception.Classification == UiRequiredExceptionClassification.ConsentRequired || exception.ErrorCode == "consent_required"
      ? GraphOperationErrorType.ConsentRequired : GraphOperationErrorType.InteractionRequired,
    "Additional authentication or consent is required for this tenant.",
    Code: exception.ErrorCode);

  public static GraphOperationError From(MsalException exception) => new(
    exception.ErrorCode is "invalid_tenant" or "tenant_not_found" or "user_not_in_tenant"
      ? GraphOperationErrorType.TenantInaccessible : GraphOperationErrorType.GraphFailure,
    "Microsoft Entra could not provide access to this tenant. Check your account access and application configuration.",
    Code: exception.ErrorCode);

  public static GraphOperationError From(ApiException exception)
  {
    var error = exception as ODataError;
    var type = exception.ResponseStatusCode switch
    {
      401 => GraphOperationErrorType.Unauthorized,
      403 => GraphOperationErrorType.Forbidden,
      404 => GraphOperationErrorType.TenantInaccessible,
      429 => GraphOperationErrorType.Throttled,
      _ => GraphOperationErrorType.GraphFailure
    };
    var message = type switch
    {
      GraphOperationErrorType.Unauthorized => "Microsoft Graph could not authenticate this request. Authenticate again.",
      GraphOperationErrorType.Forbidden => "Microsoft Graph denied access to the selected tenant.",
      GraphOperationErrorType.TenantInaccessible => "The selected tenant is not accessible through Microsoft Graph.",
      GraphOperationErrorType.Throttled => "Microsoft Graph is busy. Wait before trying again.",
      _ => "Microsoft Graph could not retrieve the organization. Try again later."
    };
    return new(type, message, exception.ResponseStatusCode, error?.Error?.Code,
      error?.Error?.InnerError?.RequestId ?? Header(exception, "request-id"),
      error?.Error?.InnerError?.ClientRequestId ?? Header(exception, "client-request-id"));
  }

  private static string? Header(ApiException exception, string name) => exception.ResponseHeaders?
    .FirstOrDefault(header => string.Equals(header.Key, name, StringComparison.OrdinalIgnoreCase)).Value?.FirstOrDefault();
}
