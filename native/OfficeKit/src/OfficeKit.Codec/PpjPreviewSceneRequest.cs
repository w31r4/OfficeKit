using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Shared by the general and PPJ-only protocol profiles, before any file parse
// or compilation. A preview scene is meaningful only for a serialized candidate.
internal static class PpjPreviewSceneRequest
{
    internal static void Validate(CodecRequest request)
    {
        if (request.PresentationProgram?.IncludePreviewScene != true) return;
        if (request.Operation != CodecOperation.CompilePpjToPptx || request.PresentationProgram.ValidationOnly)
            throw new CodecException("invalid_preview_scene_request",
                "include_preview_scene requires PPJ compilation with validation_only disabled.",
                "presentation_program.include_preview_scene");
    }
}
