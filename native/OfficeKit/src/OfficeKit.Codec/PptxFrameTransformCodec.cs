using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns only the optional rot/flipH/flipV attributes of rectangular graphic
// frames and group frames. Coordinate children remain owned by the containing
// chart, table, or group codec. Unknown attributes keep the native object out
// of the typed editable projection instead of being normalized away.
internal static class PptxFrameTransformCodec
{
    private const int MaxRotationAngle60000 = 360 * 60_000;

    internal static bool TryRead(P.Transform source, out PresentationFrameTransform? semantic)
    {
        semantic = null;
        if (source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not A.Offset offset ||
            source.ChildElements[1] is not A.Extents extents ||
            !HasOnlyAttributes(source, "rot", "flipH", "flipV") ||
            !HasOnlyAttributes(offset, "x", "y") ||
            !HasOnlyAttributes(extents, "cx", "cy") ||
            !RotationSupported(source.Rotation?.Value)) return false;
        semantic = Read(source.Rotation?.Value, source.HorizontalFlip?.Value, source.VerticalFlip?.Value);
        return true;
    }

    internal static bool TryRead(A.TransformGroup source, out PresentationFrameTransform? semantic)
    {
        semantic = null;
        if (!HasOnlyAttributes(source, "rot", "flipH", "flipV") || !RotationSupported(source.Rotation?.Value)) return false;
        semantic = Read(source.Rotation?.Value, source.HorizontalFlip?.Value, source.VerticalFlip?.Value);
        return true;
    }

    internal static void Validate(PresentationFrameTransform? transform, string elementId, string kind)
    {
        if (transform is null) return;
        if (!HasAnyField(transform))
            throw new CodecException("invalid_presentation_transform", $"Presentation {kind} {elementId} transform must define rotation or a flip.");
        if (transform.HasRotationAngle60000 && !RotationSupported(transform.RotationAngle60000))
            throw new CodecException("invalid_presentation_transform", $"Presentation {kind} {elementId} rotation must be between -360 and 360 degrees.");
    }

    internal static void Apply(P.Transform target, PresentationFrameTransform? requested)
    {
        target.Rotation = requested?.HasRotationAngle60000 == true ? requested.RotationAngle60000 : null;
        target.HorizontalFlip = requested?.HasFlipHorizontal == true ? requested.FlipHorizontal : null;
        target.VerticalFlip = requested?.HasFlipVertical == true ? requested.FlipVertical : null;
    }

    internal static void Apply(A.TransformGroup target, PresentationFrameTransform? requested)
    {
        target.Rotation = requested?.HasRotationAngle60000 == true ? requested.RotationAngle60000 : null;
        target.HorizontalFlip = requested?.HasFlipHorizontal == true ? requested.FlipHorizontal : null;
        target.VerticalFlip = requested?.HasFlipVertical == true ? requested.FlipVertical : null;
    }

    internal static void Scrub(P.Transform target)
    {
        target.Rotation = null;
        target.HorizontalFlip = null;
        target.VerticalFlip = null;
    }

    internal static void Scrub(A.TransformGroup target)
    {
        target.Rotation = null;
        target.HorizontalFlip = null;
        target.VerticalFlip = null;
    }

    private static PresentationFrameTransform? Read(int? rotation, bool? flipHorizontal, bool? flipVertical)
    {
        var semantic = new PresentationFrameTransform();
        if (rotation is { } angle) semantic.RotationAngle60000 = angle;
        if (flipHorizontal is { } horizontal) semantic.FlipHorizontal = horizontal;
        if (flipVertical is { } vertical) semantic.FlipVertical = vertical;
        return HasAnyField(semantic) ? semantic : null;
    }

    private static bool RotationSupported(int? rotation) =>
        rotation is null || Math.Abs((long)rotation.Value) <= MaxRotationAngle60000;

    private static bool HasAnyField(PresentationFrameTransform source) =>
        source.HasRotationAngle60000 || source.HasFlipHorizontal || source.HasFlipVertical;

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }
}
