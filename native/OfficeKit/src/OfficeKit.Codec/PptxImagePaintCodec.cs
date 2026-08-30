using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// One deliberately bounded DrawingML a:blipFill profile shared by native
// backgrounds and ordinary shape paint. Picture elements retain their richer
// owner codec but use the same crop/alpha/stretch-or-tile semantics.
internal static class PptxImagePaintCodec
{
    internal static bool TryRead(
        A.BlipFill? source,
        PptxPartContext? context,
        out PresentationImagePaint paint)
    {
        paint = new PresentationImagePaint();
        if (source is null || context is null || source.HasAttributes) return false;
        var children = source.ChildElements.ToArray();
        if (children.Length is < 2 or > 3 || children[0] is not A.Blip blip) return false;
        var crop = children.Length == 3 ? children[1] as A.SourceRectangle : null;
        if (children.Length == 3 && crop is null || crop is not null && !TryReadCrop(crop, out _)) return false;
        var mode = children[^1] switch
        {
            A.Stretch stretch when StretchSupported(stretch) => PresentationImagePaint.Types.Mode.Stretch,
            A.Tile tile when TileSupported(tile) => PresentationImagePaint.Types.Mode.Tile,
            _ => PresentationImagePaint.Types.Mode.Unspecified,
        };
        if (mode == PresentationImagePaint.Types.Mode.Unspecified || !TryReadBlip(blip, out var relationshipId, out var opacity))
            return false;
        try
        {
            paint.AssetId = context.ReadEmbeddedPicture(relationshipId).Id;
            paint.Mode = mode;
            if (crop is not null)
            {
                _ = TryReadCrop(crop, out var semanticCrop);
                paint.Crop = semanticCrop;
            }
            if (opacity is { } alpha) paint.OpacityThousandthPercent = alpha;
            return true;
        }
        catch (CodecException)
        {
            paint = new PresentationImagePaint();
            return false;
        }
    }

    internal static void Validate(PresentationImagePaint? source, string subject, PptxAssetCatalog? assets = null)
    {
        if (source is null) throw Invalid(subject, "image paint is missing");
        if (string.IsNullOrWhiteSpace(source.AssetId) || source.AssetId.Length > 512)
            throw Invalid(subject, "image paint asset ID must contain 1 through 512 characters");
        if (assets is not null) _ = assets.Get(source.AssetId);
        if (source.Crop is not null && !CropValuesValid(source.Crop))
            throw Invalid(subject, "image paint crop is outside the bounded signed source rectangle");
        if (source.HasOpacityThousandthPercent && source.OpacityThousandthPercent > 100_000)
            throw Invalid(subject, "image paint opacity must be between 0% and 100%");
        if (source.Mode is not (PresentationImagePaint.Types.Mode.Unspecified or
            PresentationImagePaint.Types.Mode.Stretch or PresentationImagePaint.Types.Mode.Tile))
            throw Invalid(subject, "image paint mode must be stretch or tile");
    }

    internal static A.BlipFill Build(PresentationImagePaint source, PptxPartContext context, string subject)
    {
        Validate(source, subject, context.Assets);
        var blip = new A.Blip { Embed = context.AddEmbeddedPicture(source.AssetId) };
        if (source.HasOpacityThousandthPercent)
            blip.Append(new A.AlphaModulationFixed { Amount = checked((int)source.OpacityThousandthPercent) });
        var fill = new A.BlipFill(blip);
        if (source.Crop is not null) fill.Append(BuildCrop(source.Crop));
        fill.Append(source.Mode == PresentationImagePaint.Types.Mode.Tile
            ? new A.Tile()
            : new A.Stretch(new A.FillRectangle()));
        return fill;
    }

    internal static string RelationshipId(A.BlipFill? source) =>
        source?.GetFirstChild<A.Blip>()?.Embed?.Value ?? string.Empty;

    internal static void ScrubModeledContent(A.BlipFill source)
    {
        if (source.GetFirstChild<A.Blip>() is { } blip)
        {
            blip.Embed = string.Empty;
            blip.GetFirstChild<A.AlphaModulationFixed>()?.Remove();
        }
        source.GetFirstChild<A.SourceRectangle>()?.Remove();
    }

    internal static bool TryReadCrop(A.SourceRectangle source, out PresentationImageCrop crop)
    {
        crop = new PresentationImageCrop();
        var known = new HashSet<string>(StringComparer.Ordinal) { "l", "t", "r", "b" };
        if (source.HasChildren || source.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || !known.Contains(attribute.LocalName))) return false;
        crop.LeftThousandthPercent = source.Left?.Value ?? 0;
        crop.TopThousandthPercent = source.Top?.Value ?? 0;
        crop.RightThousandthPercent = source.Right?.Value ?? 0;
        crop.BottomThousandthPercent = source.Bottom?.Value ?? 0;
        return CropValuesValid(crop);
    }

    internal static A.SourceRectangle BuildCrop(PresentationImageCrop crop) => new()
    {
        Left = crop.LeftThousandthPercent,
        Top = crop.TopThousandthPercent,
        Right = crop.RightThousandthPercent,
        Bottom = crop.BottomThousandthPercent,
    };

    internal static bool CropValuesValid(PresentationImageCrop crop) =>
        crop.LeftThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.TopThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.RightThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.BottomThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.LeftThousandthPercent + crop.RightThousandthPercent < 100_000 &&
        crop.TopThousandthPercent + crop.BottomThousandthPercent < 100_000;

    private static bool TryReadBlip(A.Blip source, out string relationshipId, out uint? opacity)
    {
        relationshipId = string.Empty;
        opacity = null;
        var attributes = source.GetAttributes().ToArray();
        if (source.Link is not null || source.CompressionState is not null || attributes.Length != 1 ||
            attributes[0].LocalName != "embed" ||
            !PptxNativeObjectCatalog.IsRelationshipNamespace(attributes[0].NamespaceUri) ||
            string.IsNullOrWhiteSpace(attributes[0].Value)) return false;
        if (source.ChildElements.Count > 1 || source.ChildElements.Any(child => child is not A.AlphaModulationFixed)) return false;
        if (source.GetFirstChild<A.AlphaModulationFixed>() is { } alpha)
        {
            if (alpha.HasChildren || alpha.Amount?.Value is not (>= 0 and <= 100_000) ||
                alpha.GetAttributes().Count != 1 || alpha.GetAttributes()[0].LocalName != "amt" ||
                alpha.GetAttributes()[0].NamespaceUri.Length != 0) return false;
            opacity = checked((uint)alpha.Amount.Value);
        }
        relationshipId = attributes[0].Value ?? string.Empty;
        return true;
    }

    private static bool StretchSupported(A.Stretch source) =>
        !source.HasAttributes && source.ChildElements.Count == 1 &&
        source.GetFirstChild<A.FillRectangle>() is { HasAttributes: false, HasChildren: false };

    private static bool TileSupported(A.Tile source) => !source.HasAttributes && !source.HasChildren;

    private static CodecException Invalid(string subject, string message) =>
        new("invalid_presentation_image_paint", $"Presentation {subject} {message}.");
}
