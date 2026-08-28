import { OfficeKitCodecError } from "./office-kit-error.mjs";

// Keep this default aligned with EffectiveCodecLimits.From in the C# codec.
// Rejecting before a defensive copy or protobuf encoding avoids turning an
// ordinary codec budget failure into a JavaScript process OOM.
export const OFFICE_KIT_DEFAULT_MAX_INPUT_BYTES = 64n * 1024n * 1024n;

export function isJavaScriptMemoryAllocationError(error) {
  if (error?.code === "ERR_BUFFER_TOO_LARGE" || error?.code === "ERR_MEMORY_ALLOCATION_FAILED") return true;
  if (!(error instanceof RangeError)) return false;
  return /(?:array buffer allocation failed|allocation failed|buffer too large|invalid (?:array|typed array|string) length|out of memory)/iu.test(String(error.message || ""));
}

export function javaScriptMemoryBudgetError(stage, cause) {
  return new OfficeKitCodecError(
    `OfficeKit JavaScript memory allocation failed during ${stage}; reduce the artifact size or use stricter codec limits.`,
    [],
    { code: "js_memory_budget_exceeded", cause },
  );
}
