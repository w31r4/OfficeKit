export function imageError(code, message, details = {}) {
  const error = new Error(message);
  error.name = "OfficeKitImageError";
  error.code = code;
  Object.assign(error, details);
  return error;
}

export function boundedImageError(error) {
  return {
    code: String(error?.code || "image-error").slice(0, 80),
    message: String(error?.message || error || "Image operation failed.").slice(0, 1_000),
  };
}
