using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeKit.Codec;

// Owns the bounded source-free media profile. Imported media remains an
// opaque/source-bound picture so an unknown timing graph is never rewritten.
internal static class PptxMediaCodec
{
    private const ulong MaxTrimMilliseconds = 86_400_000;
    private const string MediaExtensionUri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

    internal static void Validate(PresentationMedia media, string elementId, PptxAssetCatalog assets)
    {
        if (media.MediaType is not ("audio" or "video"))
            throw Invalid(elementId, "media_type must be audio or video");
        if (media.LeftEmu < 0 || media.TopEmu < 0 || media.WidthEmu <= 0 || media.HeightEmu <= 0)
            throw Invalid(elementId, "frame must use non-negative offsets and positive extents");
        if (media.HasStartAtMs && media.StartAtMs > MaxTrimMilliseconds ||
            media.HasEndAtMs && media.EndAtMs > MaxTrimMilliseconds)
            throw Invalid(elementId, $"trim offsets must be at most {MaxTrimMilliseconds} ms");
        if (media.Transform is { } transform)
        {
            if (!transform.HasRotationAngle60000 && !transform.HasFlipHorizontal && !transform.HasFlipVertical)
                throw Invalid(elementId, "transform must define rotation or a flip");
            if (transform.HasRotationAngle60000 && Math.Abs((long)transform.RotationAngle60000) > 21_600_000L)
                throw Invalid(elementId, "rotation must be between -360 and 360 degrees");
        }
        PptxNonVisualAccessibilityCodec.Validate(media.Accessibility, elementId, "media");

        var source = assets.GetMedia(media.AssetId);
        var supported = media.MediaType == "video"
            ? source.ContentType.Equals("video/mp4", StringComparison.OrdinalIgnoreCase)
            : source.ContentType is "audio/mpeg" or "audio/mp4" or "audio/wav" or "audio/x-wav";
        if (!supported)
            throw Invalid(elementId, $"{media.MediaType} cannot use asset content type {source.ContentType}");
        _ = assets.Get(media.PosterAssetId);
    }

    internal static P.Picture Build(
        PresentationElement source,
        uint nativeId,
        PptxPartContext context,
        SlidePart slidePart,
        PresentationDocument package)
    {
        var media = source.Media;
        var assets = context.Assets ?? throw Invalid(source.Id, "authoring requires an asset catalog");
        Validate(media, source.Id, assets);
        var mediaAsset = assets.GetMedia(media.AssetId);
        var mediaPart = assets.GetOrCreateMediaPart(package, media.AssetId);
        var mediaRelationshipId = NextRelationshipId(slidePart, "Media", mediaAsset.Sha256);
        slidePart.AddMediaReferenceRelationship(mediaPart, mediaRelationshipId);
        var fileRelationshipId = NextRelationshipId(slidePart, media.MediaType == "video" ? "Video" : "Audio", mediaAsset.Sha256);
        if (media.MediaType == "video") slidePart.AddVideoReferenceRelationship(mediaPart, fileRelationshipId);
        else slidePart.AddAudioReferenceRelationship(mediaPart, fileRelationshipId);

        var nonVisualDrawing = new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisualDrawing, media.Accessibility);
        nonVisualDrawing.Append(new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" });

        var mediaExtension = new P14.Media { Embed = mediaRelationshipId };
        if (media.HasStartAtMs || media.HasEndAtMs)
        {
            var trim = new P14.MediaTrim();
            if (media.HasStartAtMs) trim.Start = media.StartAtMs.ToString(CultureInfo.InvariantCulture);
            if (media.HasEndAtMs) trim.End = media.EndAtMs.ToString(CultureInfo.InvariantCulture);
            mediaExtension.Append(trim);
        }
        var applicationProperties = new P.ApplicationNonVisualDrawingProperties();
        if (media.MediaType == "video") applicationProperties.Append(new A.VideoFromFile { Link = fileRelationshipId });
        else applicationProperties.Append(new A.AudioFromFile { Link = fileRelationshipId });
        applicationProperties.Append(new P.ApplicationNonVisualDrawingPropertiesExtensionList(
            new P.ApplicationNonVisualDrawingPropertiesExtension(mediaExtension) { Uri = MediaExtensionUri }));

        var fill = new P.BlipFill(
            new A.Blip { Embed = context.AddEmbeddedPicture(media.PosterAssetId) },
            new A.Stretch(new A.FillRectangle()));
        var transform = new A.Transform2D(
            new A.Offset { X = media.LeftEmu, Y = media.TopEmu },
            new A.Extents { Cx = media.WidthEmu, Cy = media.HeightEmu });
        ApplyTransform(transform, media.Transform);
        return new P.Picture(
            new P.NonVisualPictureProperties(
                nonVisualDrawing,
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                applicationProperties),
            fill,
            new P.ShapeProperties(
                transform,
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }

    private static void ApplyTransform(A.Transform2D target, PresentationImageTransform? source)
    {
        if (source is null) return;
        target.Rotation = source.HasRotationAngle60000 ? source.RotationAngle60000 : null;
        target.HorizontalFlip = source.HasFlipHorizontal ? source.FlipHorizontal : null;
        target.VerticalFlip = source.HasFlipVertical ? source.FlipVertical : null;
    }

    private static string NextRelationshipId(SlidePart slidePart, string kind, string sha256)
    {
        var used = slidePart.Parts.Select(pair => pair.RelationshipId)
            .Concat(slidePart.ExternalRelationships.Select(relationship => relationship.Id))
            .Concat(slidePart.HyperlinkRelationships.Select(relationship => relationship.Id))
            .Concat(slidePart.DataPartReferenceRelationships.Select(relationship => relationship.Id))
            .ToHashSet(StringComparer.Ordinal);
        var digest = sha256[..16].ToLowerInvariant();
        var stem = $"rIdOfficeKit{kind}{digest}_";
        for (var index = 1; index <= 1_000_000; index++)
        {
            var candidate = stem + index.ToString(CultureInfo.InvariantCulture);
            if (!used.Contains(candidate)) return candidate;
        }
        throw new CodecException(
            "presentation_relationship_budget_exceeded",
            $"PPTX {kind.ToLowerInvariant()} relationship ID allocation exceeded its bounded search.");
    }

    private static CodecException Invalid(string elementId, string message) =>
        new("invalid_presentation_media", $"Presentation media {elementId} {message}.");
}
