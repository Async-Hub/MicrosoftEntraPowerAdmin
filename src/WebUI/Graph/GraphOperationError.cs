namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Graph;

public enum GraphOperationErrorType
{
  TenantNotSelected,
  InteractionRequired,
  ConsentRequired,
  Unauthorized,
  Forbidden,
  TenantInaccessible,
  TenantMismatch,
  TenantChanged,
  Throttled,
  GraphFailure,
  InvalidInput,
  NotFound,
  StaleState,
  ConfirmationRequired
}

public sealed record GraphOperationError(
  GraphOperationErrorType Type,
  string Message,
  int? HttpStatus = null,
  string? Code = null,
  string? RequestId = null,
  string? ClientRequestId = null);
