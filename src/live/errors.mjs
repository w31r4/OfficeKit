export class OfficeLiveError extends Error {
  constructor(code, message, { retryable = false, maybeApplied = false, details } = {}) {
    super(message);
    this.name = "OfficeLiveError";
    this.code = code;
    this.retryable = retryable;
    this.maybeApplied = maybeApplied;
    if (details !== undefined) this.details = details;
  }
}

export function officeLiveError(code, message, options) {
  return new OfficeLiveError(code, message, options);
}

export function toOfficeLiveFailure(error) {
  if (error?.code) return error;
  return officeLiveError("internal-error", error instanceof Error ? error.message : String(error));
}
