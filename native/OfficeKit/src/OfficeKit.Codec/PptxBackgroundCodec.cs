using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns only a direct, bounded p:bg choice on one p:cSld. Effective color
// inheritance stays in the JavaScript model. Literal RGB linear gradients and
// one centered radial profile are typed; patterns, transforms and other
// effect-bearing backgrounds remain source-bound. Images use the shared
// embedded crop/alpha/stretch-or-default-tile profile.
internal static class PptxBackgroundCodec
{
    internal static PresentationBackground? Read(P.CommonSlideData? source, PptxPartContext? context = null)
    {
        var background = source?.GetFirstChild<P.Background>();
        return background is not null && TryRead(background, context, out var semantic) ? semantic : null;
    }

    internal static bool Supports(P.CommonSlideData? source, PptxPartContext? context = null)
    {
        var backgrounds = source?.Elements<P.Background>().ToArray() ?? [];
        return backgrounds.Length == 0 || backgrounds.Length == 1 && TryRead(backgrounds[0], context, out _);
    }

    internal static void Validate(PresentationBackground? source, PptxAssetCatalog? assets = null)
    {
        if (source is null) return;
        if (source.GradientFill is not null)
        {
            if (!string.IsNullOrWhiteSpace(source.ImageAssetId) ||
                source.ImagePaint is not null ||
                source.ColorCase != PresentationBackground.ColorOneofCase.None ||
                source.KindCase != PresentationBackground.KindOneofCase.None ||
                source.ImageAlphaModulationFixed)
                throw Invalid("Presentation gradient background cannot also define image, color, kind, or image effects.");
            PptxGradientFillCodec.Validate(source.GradientFill, "Presentation background");
            return;
        }
        if (source.ImagePaint is not null)
        {
            if (!string.IsNullOrWhiteSpace(source.ImageAssetId) || source.ImageAlphaModulationFixed ||
                source.ColorCase != PresentationBackground.ColorOneofCase.None ||
                source.KindCase != PresentationBackground.KindOneofCase.None)
                throw Invalid("Presentation canonical image background cannot also define legacy image, color, kind, or image effects.");
            PptxImagePaintCodec.Validate(source.ImagePaint, "background", assets);
            return;
        }
        if (!string.IsNullOrWhiteSpace(source.ImageAssetId))
        {
            if (source.ColorCase != PresentationBackground.ColorOneofCase.None || source.KindCase != PresentationBackground.KindOneofCase.None)
                throw Invalid("Presentation image background cannot also define a color or kind.");
            if (source.ImageAssetId.Length > 512)
                throw Invalid("Presentation image background asset ID exceeds 512 characters.");
            if (assets is not null) _ = assets.Get(source.ImageAssetId);
            return;
        }
        if (source.ImageAlphaModulationFixed)
            throw Invalid("Presentation image alpha modulation requires an image background asset.");
        switch (source.ColorCase)
        {
            case PresentationBackground.ColorOneofCase.ColorRgb:
                _ = PptxColor.Normalize(source.ColorRgb);
                break;
            case PresentationBackground.ColorOneofCase.ColorScheme:
                _ = PptxColor.NormalizeScheme(source.ColorScheme);
                break;
            default:
                throw Invalid("Presentation background requires exactly one RGB or theme color.");
        }
        switch (source.KindCase)
        {
            case PresentationBackground.KindOneofCase.Solid when source.Solid:
                break;
            case PresentationBackground.KindOneofCase.StyleReferenceIndex:
                break;
            default:
                throw Invalid("Presentation background requires a solid mode or style-reference index.");
        }
    }

    internal static void Build(P.CommonSlideData target, PresentationBackground? source, PptxPartContext context)
    {
        if (source is null) return;
        target.AddChild(BuildElement(source, context), true);
    }

    internal static void Apply(P.CommonSlideData target, PresentationBackground? source, PptxPartContext context)
    {
        Validate(source, context.Assets);
        var current = target.GetFirstChild<P.Background>();
        var currentImageRelationshipId = ImageRelationshipId(current);
        if (source is null)
        {
            current?.Remove();
            context.RemoveIfUnreferenced(currentImageRelationshipId);
            return;
        }
        var replacement = BuildElement(source, context);
        var replacementImageRelationshipId = ImageRelationshipId(replacement);
        if (current is null)
        {
            target.AddChild(replacement, true);
            context.RemoveIfUnreferenced(currentImageRelationshipId);
            return;
        }
        current.InsertAfterSelf(replacement);
        current.Remove();
        context.RemoveIfUnreferenced(currentImageRelationshipId == replacementImageRelationshipId ? string.Empty : currentImageRelationshipId);
    }

    internal static void ScrubModeledContent(P.CommonSlideData? source, PptxPartContext? context = null)
    {
        var background = source?.GetFirstChild<P.Background>();
        if (background is not null && TryRead(background, context, out _)) background.Remove();
    }

    private static bool TryRead(P.Background source, PptxPartContext? context, out PresentationBackground semantic)
    {
        semantic = new PresentationBackground();
        if (source.GetAttributes().Count != 0 || source.ChildElements.Count != 1) return false;
        switch (source.FirstChild)
        {
            case P.BackgroundProperties properties:
                if (properties.GetAttributes().Count != 0) return false;
                var children = properties.ChildElements.ToArray();
                if (children.Length == 1 && children[0] is A.BlipFill imageFill)
                {
                    if (PptxImagePaintCodec.TryRead(imageFill, context, out var imagePaint))
                    {
                        semantic.ImagePaint = imagePaint;
                        return true;
                    }
                    if (!TryReadImage(imageFill, context, out var assetId, out var alphaModulationFixed)) return false;
                    semantic.ImageAssetId = assetId;
                    semantic.ImageAlphaModulationFixed = alphaModulationFixed;
                    return true;
                }
                if (children.Length is < 1 or > 2) return false;
                if (children.Length == 2 && (children[1] is not A.EffectList { ChildElements.Count: 0 } effectList || effectList.GetAttributes().Count != 0)) return false;
                if (children[0] is A.GradientFill gradient)
                {
                    if (!PptxGradientFillCodec.TryRead(gradient, out var gradientSemantic)) return false;
                    semantic.GradientFill = gradientSemantic;
                    return true;
                }
                if (children[0] is not A.SolidFill solid) return false;
                if (!TryReadColor(solid, out semantic)) return false;
                semantic.Solid = true;
                return true;
            case P.BackgroundStyleReference reference:
                if (reference.Index?.Value is not { } index ||
                    reference.GetAttributes().Any(attribute => attribute.LocalName != "idx") ||
                    !TryReadColor(reference, out semantic)) return false;
                semantic.StyleReferenceIndex = index;
                return true;
            default:
                return false;
        }
    }

    private static bool TryReadImage(A.BlipFill source, PptxPartContext? context, out string assetId, out bool alphaModulationFixed)
    {
        assetId = string.Empty;
        alphaModulationFixed = false;
        if (context is null || context.Assets is null || source.GetAttributes().Count != 0 || source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not A.Blip blip || source.ChildElements[1] is not A.Stretch stretch ||
            stretch.GetAttributes().Count != 0 || stretch.ChildElements.Count != 1 ||
            stretch.GetFirstChild<A.FillRectangle>() is not { } fillRect || fillRect.GetAttributes().Count != 0 || fillRect.ChildElements.Count != 0 ||
            blip.Link is not null || blip.Embed?.Value is not { Length: > 0 } embed || blip.CompressionState is not null ||
            blip.ChildElements.Any(child => child is not A.AlphaModulationFixed || child.GetAttributes().Count != 0 || child.ChildElements.Count != 0) ||
            blip.ChildElements.Count > 1 || blip.GetAttributes().Count != 1)
            return false;
        var attribute = blip.GetAttributes()[0];
        if (attribute.LocalName != "embed" ||
            attribute.NamespaceUri is not "http://schemas.openxmlformats.org/officeDocument/2006/relationships" and
            not "http://purl.oclc.org/ooxml/officeDocument/relationships") return false;
        try
        {
            assetId = context.ReadEmbeddedPicture(embed).Id;
            alphaModulationFixed = blip.GetFirstChild<A.AlphaModulationFixed>() is not null;
            return true;
        }
        catch (CodecException)
        {
            assetId = string.Empty;
            alphaModulationFixed = false;
            return false;
        }
    }

    private static bool TryReadColor(OpenXmlCompositeElement source, out PresentationBackground target)
    {
        target = new PresentationBackground();
        if ((source is A.SolidFill && source.GetAttributes().Count != 0) || source.ChildElements.Count != 1) return false;
        switch (source.FirstChild)
        {
            case A.RgbColorModelHex rgb when rgb.ChildElements.Count == 0 && rgb.GetAttributes().All(attribute => attribute.LocalName == "val") && rgb.Val?.Value is { Length: 6 } value:
                try { target.ColorRgb = PptxColor.Normalize(value); return true; }
                catch (CodecException) { return false; }
            case A.SchemeColor scheme when scheme.ChildElements.Count == 0 && scheme.GetAttributes().All(attribute => attribute.LocalName == "val") && scheme.Val?.Value is { } value && PptxColor.TrySchemeToken(value, out var token):
                target.ColorScheme = token;
                return true;
            default:
                return false;
        }
    }

    private static string ImageRelationshipId(P.Background? source) =>
        source?.BackgroundProperties?.GetFirstChild<A.BlipFill>()?.GetFirstChild<A.Blip>()?.Embed?.Value ?? string.Empty;

    private static P.Background BuildElement(PresentationBackground source, PptxPartContext context)
    {
        Validate(source, context.Assets);
        if (source.ImagePaint is not null)
            return new P.Background(new P.BackgroundProperties(
                PptxImagePaintCodec.Build(source.ImagePaint, context, "background")));
        if (!string.IsNullOrWhiteSpace(source.ImageAssetId))
        {
            var blip = new A.Blip { Embed = context.AddEmbeddedPicture(source.ImageAssetId) };
            if (source.ImageAlphaModulationFixed) blip.Append(new A.AlphaModulationFixed());
            var fill = new A.BlipFill(blip);
            fill.Append(new A.Stretch(new A.FillRectangle()));
            return new P.Background(new P.BackgroundProperties(fill));
        }
        if (source.GradientFill is not null)
            return new P.Background(new P.BackgroundProperties(
                PptxGradientFillCodec.Build(source.GradientFill, "Presentation background"),
                new A.EffectList()));
        return source.KindCase switch
        {
            PresentationBackground.KindOneofCase.Solid =>
                new P.Background(new P.BackgroundProperties(new A.SolidFill(Color(source)), new A.EffectList())),
            PresentationBackground.KindOneofCase.StyleReferenceIndex =>
                new P.Background(new P.BackgroundStyleReference(Color(source)) { Index = source.StyleReferenceIndex }),
            _ => throw Invalid("Presentation background kind is missing."),
        };
    }

    private static OpenXmlElement Color(PresentationBackground source) => source.ColorCase switch
    {
        PresentationBackground.ColorOneofCase.ColorRgb => new A.RgbColorModelHex { Val = PptxColor.Normalize(source.ColorRgb) },
        PresentationBackground.ColorOneofCase.ColorScheme => new A.SchemeColor { Val = PptxColor.SchemeValue(source.ColorScheme) },
        _ => throw Invalid("Presentation background color is missing."),
    };

    private static CodecException Invalid(string message) => new("invalid_presentation_background", message);
}
