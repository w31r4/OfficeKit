export class OfficeKitCodecError extends Error {
  constructor(message, diagnostics = [], options = {}) {
    super(message, options);
    this.name = "OfficeKitCodecError";
    this.code = diagnostics[0]?.code || options.code || "office_kit_codec_error";
    this.diagnostics = diagnostics;
  }
}
