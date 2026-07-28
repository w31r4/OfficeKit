export class ExcelLiveError extends Error {
  constructor(code, message, { retryable = false, maybeApplied = false, details } = {}) {
    super(message);
    this.name = "ExcelLiveError";
    this.code = code;
    this.retryable = retryable;
    this.maybeApplied = maybeApplied;
    if (details !== undefined) this.details = details;
  }
}

export function excelLiveError(code, message, options) {
  return new ExcelLiveError(code, message, options);
}

export function toExcelLiveFailure(error) {
  if (error instanceof ExcelLiveError) {
    return {
      code: error.code,
      message: error.message,
      retryable: error.retryable,
      maybeApplied: error.maybeApplied,
      ...(error.details === undefined ? {} : { details: error.details }),
    };
  }
  return {
    code: "internal-error",
    message: error instanceof Error ? error.message : String(error),
    retryable: false,
    maybeApplied: false,
  };
}
