namespace AsyncHub.MicrosoftEntraPowerAdmin.WebUI.Infrastructure;

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
  ConfirmationRequired,
  PolicyIsAssigned
}

public sealed record GraphOperationError(
  GraphOperationErrorType Type,
  string Message,
  int? HttpStatus = null,
  string? Code = null,
  string? RequestId = null,
  string? ClientRequestId = null);
