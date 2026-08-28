using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the deliberately bounded, source-preserving p:pic projection. The
// semantic image exposes only the asset/frame/crop/accessibility leaves; the
// standard DrawingML effects, SVG fallback and presentation extensions below
// are accepted as preserved source structure, never reconstructed by the
// authored-image builder.
internal static class PptxPictureCodec
{
    private const int MaxTextLength = 1_024;
    // A source-bound image may legitimately bleed past the slide edge. Keep
    // that valid DrawingML geometry bounded so malformed imports cannot turn
    // into unbounded JS coordinates or allocations.
    private const long MaxFrameCoordinateEmu = 100_000_000L;
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string Office2010Namespace = "http://schemas.microsoft.com/office/drawing/2010/main";
    private const string SvgBlipNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private const string UseLocalDpiUri = "{28A0092B-C50C-407E-A947-70E740481C1C}";
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    private const string HiddenFillUri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}";

    internal static bool TryRead(P.Picture source, PptxPartContext context, out PresentationImage image)
    {
        image = new PresentationImage();
        if (!TryParts(source, out var nonVisual, out var blip, out var transform, out var crop)) return false;
        // Accessibility is a leaf capability, not ownership of the whole
        // picture. An ambiguous known extension hides the modeled metadata and
        // disables only that setter; unrelated picture edits remain residual-
        // preserving and source-bound.
        _ = PptxNonVisualAccessibilityCodec.TryReadResidual(nonVisual, out var accessibility);
        var relationshipId = blip.Embed?.Value ?? string.Empty;
        if (relationshipId.Length == 0) return false;
        try
        {
            var asset = context.ReadEmbeddedPicture(relationshipId);
            var offset = transform.Offset!;
            var extents = transform.Extents!;
            image = new PresentationImage
            {
                AssetId = asset.Id,
                AltText = accessibility?.HasDescription == true ? accessibility.Description : string.Empty,
                AccessibilityTitle = accessibility?.HasTitle == true ? accessibility.Title : string.Empty,
                LeftEmu = offset.X?.Value ?? 0,
                TopEmu = offset.Y?.Value ?? 0,
                WidthEmu = extents.Cx?.Value ?? 0,
                HeightEmu = extents.Cy?.Value ?? 0,
            };
            if (accessibility?.HasDecorative == true) image.AccessibilityDecorative = accessibility.Decorative;
            if (crop is not null) image.Crop = ReadCrop(crop);
            var visual = ReadTransform(transform);
            if (visual is not null) image.Transform = visual;
            PptxNonVisualAccessibilityCodec.Validate(Accessibility(image), "source", "image");
            return FrameCoordinateSupported(image.LeftEmu, allowNegative: true) &&
                   FrameCoordinateSupported(image.TopEmu, allowNegative: true) &&
                   FrameExtentSupported(image.WidthEmu) && FrameExtentSupported(image.HeightEmu) &&
                   (nonVisual.Name?.Value?.Length ?? 0) <= MaxTextLength;
        }
        catch (CodecException)
        {
            image = new PresentationImage();
            return false;
        }
    }

    internal static void Validate(PresentationImage? image, string elementId, PptxAssetCatalog assets, bool sourceBound = false)
    {
        if (image is null)
            throw Invalid(elementId, "payload is missing");
        if (string.IsNullOrWhiteSpace(image.AssetId) || image.AssetId.Length > 512)
            throw Invalid(elementId, "asset ID must contain 1 through 512 characters");
        _ = assets.Get(image.AssetId);
        PptxNonVisualAccessibilityCodec.Validate(Accessibility(image), elementId, "image");
        if (!FrameCoordinateSupported(image.LeftEmu, sourceBound) || !FrameCoordinateSupported(image.TopEmu, sourceBound) ||
            !FrameExtentSupported(image.WidthEmu) || !FrameExtentSupported(image.HeightEmu))
            throw Invalid(elementId, sourceBound
                ? "source-bound frame coordinates must stay within the bounded DrawingML range and extents must be positive"
                : "frame must use non-negative coordinates and positive extents");
        if (image.Crop is not null && !CropValuesValid(image.Crop))
            throw Invalid(elementId, "crop edges must be between -100% and 100% and opposing sums must remain below 100%");
        ValidateTransform(image.Transform, elementId);
    }

    internal static P.Picture Build(PresentationElement source, uint nativeId, PptxPartContext context)
    {
        var image = source.Image;
        var transform = new A.Transform2D(
            new A.Offset { X = image.LeftEmu, Y = image.TopEmu },
            new A.Extents { Cx = image.WidthEmu, Cy = image.HeightEmu });
        ApplyTransform(transform, image.Transform);
        var fill = new P.BlipFill(new A.Blip { Embed = context.AddEmbeddedPicture(image.AssetId) });
        if (image.Crop is not null) fill.Append(BuildCrop(image.Crop));
        fill.Append(new A.Stretch(new A.FillRectangle()));
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, Accessibility(image));
        return new P.Picture(
            new P.NonVisualPictureProperties(
                nonVisual,
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            fill,
            new P.ShapeProperties(
                transform,
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }

    internal static void Apply(P.Picture source, PresentationElement requested, PptxPartContext context)
    {
        if (!TryRead(source, context, out var currentImage) ||
            !TryParts(source, out var nonVisual, out var blip, out var transform, out _))
            throw new CodecException("unsupported_presentation_edit", $"Presentation image {requested.Id} no longer matches the editable picture profile.");
        var current = context.ReadEmbeddedPicture(blip.Embed?.Value ?? string.Empty);
        var replacement = context.Assets?.Get(requested.Image.AssetId) ??
            throw new CodecException("invalid_presentation_asset", $"Presentation image {requested.Id} requires an asset catalog.");
        if (!current.Id.Equals(replacement.Id, StringComparison.Ordinal))
        {
            if (!current.ContentType.Equals(replacement.ContentType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} replacement must retain content type {current.ContentType}.");
            if (HasSvgFallback(blip))
                throw new CodecException(
                    "unsupported_presentation_image",
                    $"Presentation image {requested.Id} has a paired SVG fallback; replacing only the primary asset would leave the fallback stale.");
            blip.Embed = context.AddEmbeddedPicture(replacement.Id);
        }
        nonVisual.Name = requested.Name;
        if (!AccessibilityEqual(currentImage, requested.Image))
            PptxNonVisualAccessibilityCodec.ApplyResidualBound(nonVisual, Accessibility(requested.Image), "image");
        transform.Offset!.X = requested.Image.LeftEmu;
        transform.Offset.Y = requested.Image.TopEmu;
        transform.Extents!.Cx = requested.Image.WidthEmu;
        transform.Extents.Cy = requested.Image.HeightEmu;
        ApplyCrop(source.BlipFill!, requested.Image.Crop);
        ApplyTransform(transform, requested.Image.Transform);
    }

    internal static void ScrubModeledContent(P.Picture source)
    {
        if (source.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            nonVisual.Name = string.Empty;
            PptxNonVisualAccessibilityCodec.ScrubResidualModeledContent(nonVisual);
            nonVisual.Description = string.Empty;
        }
        if (source.BlipFill?.GetFirstChild<A.Blip>() is { } blip) blip.Embed = string.Empty;
        source.BlipFill?.GetFirstChild<A.SourceRectangle>()?.Remove();
        if (source.ShapeProperties?.Transform2D is { } transform)
        {
            if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
            if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
            transform.Rotation = null;
            transform.HorizontalFlip = null;
            transform.VerticalFlip = null;
        }
    }

    private static bool TryParts(
        P.Picture source,
        out P.NonVisualDrawingProperties nonVisual,
        out A.Blip blip,
        out A.Transform2D transform,
        out A.SourceRectangle? crop)
    {
        nonVisual = null!;
        blip = null!;
        transform = null!;
        crop = null;
        var nonVisualContainer = source.NonVisualPictureProperties;
        var fill = source.BlipFill;
        var properties = source.ShapeProperties;
        if (nonVisualContainer?.NonVisualDrawingProperties is not { } nv ||
            nonVisualContainer.NonVisualPictureDrawingProperties is null ||
            nonVisualContainer.ApplicationNonVisualDrawingProperties is null ||
            fill is null || properties is null) return false;
        var fillChildren = fill.ChildElements.ToArray();
        if (source.ChildElements.Count != 3 ||
            nonVisualContainer.ChildElements.Count != 3 ||
            fillChildren.Length is < 2 or > 3 || fillChildren[0] is not A.Blip embedded ||
            fillChildren[^1] is not A.Stretch stretch ||
            fillChildren.Length == 3 && fillChildren[1] is not A.SourceRectangle ||
            !StretchSupported(stretch) ||
            !BlipSupported(embedded) || !BlipFillSupported(fill) ||
            !ShapePropertiesSupported(properties) ||
            properties.Elements<A.Transform2D>().SingleOrDefault() is not { } xfrm ||
            properties.Elements<A.PresetGeometry>().SingleOrDefault() is not { } geometry ||
            geometry.Preset is null || geometry.GetAttributes().Count != 1 ||
            geometry.GetAttributes()[0].LocalName != "prst" || geometry.GetAttributes()[0].NamespaceUri.Length != 0 ||
            geometry.ChildElements.Count != 1 || geometry.GetFirstChild<A.AdjustValueList>() is not { } adjustments ||
            adjustments.HasAttributes || adjustments.HasChildren ||
            !TransformSupported(xfrm)) return false;
        crop = fillChildren.Length == 3 ? (A.SourceRectangle)fillChildren[1] : null;
        if (crop is not null && !CropSupported(crop)) return false;
        nonVisual = nv;
        blip = embedded;
        transform = xfrm;
        return true;
    }

    private static bool BlipSupported(A.Blip blip)
    {
        var attributes = blip.GetAttributes().ToArray();
        var embed = attributes.Where(attribute => attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri)).ToArray();
        var link = attributes.Where(attribute => attribute.LocalName == "link" && RelationshipNamespace(attribute.NamespaceUri)).ToArray();
        if (embed.Length != 1 || link.Length != 0 || string.IsNullOrWhiteSpace(embed[0].Value)) return false;
        if (attributes.Any(attribute =>
                !((attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri)) ||
                  (attribute.LocalName == "cstate" && attribute.NamespaceUri.Length == 0)))) return false;
        if (attributes.Any(attribute => attribute.LocalName == "cstate" && attribute.Value is not ("email" or "screen" or "print" or "hqprint"))) return false;
        return BlipChildrenSupported(blip);
    }

    private static bool BlipChildrenSupported(A.Blip blip)
    {
        var extensionLists = blip.ChildElements.Where(child => child.LocalName == "extLst" && child.NamespaceUri == DrawingNamespace).ToArray();
        if (blip.ChildElements.Any(child => child.LocalName is not ("alphaModFix" or "clrChange" or "extLst") || child.NamespaceUri != DrawingNamespace) ||
            extensionLists.Length > 1) return false;
        var alpha = blip.ChildElements.Where(child => child.LocalName == "alphaModFix").ToArray();
        var colors = blip.ChildElements.Where(child => child.LocalName == "clrChange").ToArray();
        return alpha.Length <= 1 && colors.Length <= 1 &&
               alpha.All(AlphaModulationSupported) && colors.All(ColorChangeSupported) &&
               extensionLists.All(BlipExtensionListSupported);
    }

    private static bool AlphaModulationSupported(OpenXmlElement element)
    {
        var attributes = element.GetAttributes().ToArray();
        return element.NamespaceUri == DrawingNamespace && !element.HasChildren &&
               attributes.All(attribute => attribute.LocalName == "amt" && attribute.NamespaceUri.Length == 0) &&
               attributes.Length <= 1 && attributes.All(attribute => UIntInRange(attribute.Value ?? string.Empty, 0, 100_000));
    }

    private static bool ColorChangeSupported(OpenXmlElement element)
    {
        if (element.NamespaceUri != DrawingNamespace || element.HasAttributes || element.ChildElements.Count != 2) return false;
        var from = element.ChildElements.FirstOrDefault(child => child.LocalName == "clrFrom" && child.NamespaceUri == DrawingNamespace);
        var to = element.ChildElements.FirstOrDefault(child => child.LocalName == "clrTo" && child.NamespaceUri == DrawingNamespace);
        return from is not null && to is not null && element.ChildElements.All(child => ReferenceEquals(child, from) || ReferenceEquals(child, to)) &&
               ColorChoiceSupported(from, allowAlpha: false) && ColorChoiceSupported(to, allowAlpha: true);
    }

    private static bool ColorChoiceSupported(OpenXmlElement element, bool allowAlpha)
    {
        if (element.HasAttributes || element.ChildElements.Count != 1 || element.ChildElements[0].LocalName != "srgbClr" ||
            element.ChildElements[0].NamespaceUri != DrawingNamespace) return false;
        var color = element.ChildElements[0];
        var attributes = color.GetAttributes().ToArray();
        if (attributes.Any(attribute => attribute.LocalName != "val" || attribute.NamespaceUri.Length != 0) ||
            attributes.Length != 1 || !IsHexColor(attributes[0].Value ?? string.Empty)) return false;
        var alpha = color.ChildElements.ToArray();
        return alpha.Length == 0 || allowAlpha && alpha.Length == 1 && alpha[0].LocalName == "alpha" &&
               alpha[0].NamespaceUri == DrawingNamespace && alpha[0].ChildElements.Count == 0 &&
               alpha[0].GetAttributes().Count == 1 && alpha[0].GetAttributes()[0].LocalName == "val" &&
               alpha[0].GetAttributes()[0].NamespaceUri.Length == 0 && UIntInRange(alpha[0].GetAttributes()[0].Value ?? string.Empty, 0, 100_000);
    }

    private static bool BlipExtensionListSupported(OpenXmlElement element)
    {
        if (element is not A.BlipExtensionList || element.NamespaceUri != DrawingNamespace) return false;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var extension in element.ChildElements)
        {
            if (extension.LocalName != "ext" || extension.NamespaceUri != DrawingNamespace || extension.GetAttributes().Count != 1 ||
                extension.GetAttributes()[0].LocalName != "uri" || extension.GetAttributes()[0].NamespaceUri.Length != 0 ||
                !seen.Add(extension.GetAttributes()[0].Value ?? string.Empty) || extension.ChildElements.Count != 1) return false;
            var uri = extension.GetAttributes()[0].Value ?? string.Empty;
            if (uri == UseLocalDpiUri && !UseLocalDpiSupported(extension.ChildElements[0])) return false;
            if (uri == SvgExtensionUri && !SvgFallbackSupported(extension.ChildElements[0])) return false;
            if (uri is not (UseLocalDpiUri or SvgExtensionUri)) return false;
        }
        return true;
    }

    private static bool UseLocalDpiSupported(OpenXmlElement element)
    {
        var attributes = element.GetAttributes().ToArray();
        return element.LocalName == "useLocalDpi" && element.NamespaceUri == Office2010Namespace && !element.HasChildren &&
               attributes.All(attribute => attribute.LocalName == "val" && attribute.NamespaceUri.Length == 0 && attribute.Value is "0" or "1") &&
               attributes.Length <= 1;
    }

    private static bool SvgFallbackSupported(OpenXmlElement element)
    {
        var attributes = element.GetAttributes().ToArray();
        var embeds = attributes.Where(attribute => attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri)).ToArray();
        return element.LocalName == "svgBlip" && element.NamespaceUri == SvgBlipNamespace && !element.HasChildren && embeds.Length == 1 &&
               attributes.Length == 1 && !string.IsNullOrWhiteSpace(embeds[0].Value);
    }

    private static bool HasSvgFallback(A.Blip blip) => blip.ChildElements
        .Where(child => child.LocalName == "extLst" && child.NamespaceUri == DrawingNamespace)
        .SelectMany(child => child.ChildElements)
        .Any(child => child.LocalName == "ext" && child.GetAttributes().Any(attribute => attribute.LocalName == "uri" && attribute.Value == SvgExtensionUri));

    private static bool BlipFillSupported(P.BlipFill fill)
    {
        var attributes = fill.GetAttributes().ToArray();
        return attributes.All(attribute => attribute.LocalName == "rotWithShape" && attribute.NamespaceUri.Length == 0 && attribute.Value is "0" or "1");
    }

    private static bool StretchSupported(A.Stretch stretch)
    {
        if (stretch.HasAttributes || stretch.ChildElements.Count > 1) return false;
        return stretch.ChildElements.Count == 0 || stretch.GetFirstChild<A.FillRectangle>() is { } fillRect && !fillRect.HasAttributes && !fillRect.HasChildren;
    }

    private static bool ShapePropertiesSupported(P.ShapeProperties properties)
    {
        var attributes = properties.GetAttributes().ToArray();
        if (attributes.Any(attribute => attribute.LocalName != "bwMode" || attribute.NamespaceUri.Length != 0 || attribute.Value is not ("auto" or "gray" or "ltGray" or "invGray" or "grayWhite" or "blackGray" or "blackWhite" or "clr"))) return false;
        if (properties.ChildElements.Count < 2 || properties.ChildElements.Count > 6) return false;
        if (properties.ChildElements.Count(child => child is A.Transform2D) != 1 || properties.ChildElements.Count(child => child is A.PresetGeometry) != 1) return false;
        return properties.ChildElements.All(child => child switch
        {
            A.Transform2D or A.PresetGeometry => true,
            A.NoFill noFill => !noFill.HasAttributes && !noFill.HasChildren,
            A.Outline outline => LineSupported(outline),
            A.EffectList effects => effects.HasAttributes || effects.HasChildren, // preserved, never edited by the image facade
            A.ShapePropertiesExtensionList extensions => ShapeExtensionListSupported(extensions),
            _ => false,
        });
    }

    private static bool LineSupported(A.Outline outline)
    {
        var attributes = outline.GetAttributes().ToArray();
        if (attributes.Any(attribute => attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("w" or "cap" or "cmpd" or "algn"))) return false;
        return outline.ChildElements.All(child => child.LocalName is "noFill" or "solidFill" or "prstDash" or "round" or "bevel" or "miter" or "headEnd" or "tailEnd") &&
               outline.ChildElements.All(child => child.LocalName != "noFill" || !child.HasAttributes && !child.HasChildren);
    }

    private static bool ShapeExtensionListSupported(A.ShapePropertiesExtensionList extensions)
    {
        if (extensions.HasAttributes) return false;
        foreach (var extension in extensions.ChildElements)
        {
            var attributes = extension.GetAttributes().ToArray();
            if (extension.LocalName != "ext" || extension.NamespaceUri != DrawingNamespace || attributes.Length != 1 ||
                attributes[0].LocalName != "uri" || attributes[0].NamespaceUri.Length != 0 || attributes[0].Value != HiddenFillUri ||
                extension.ChildElements.Count != 1) return false;
            var hidden = extension.ChildElements[0];
            if (hidden.LocalName != "hiddenFill" || hidden.NamespaceUri != Office2010Namespace || hidden.ChildElements.Count != 1 || hidden.GetAttributes().Count != 0)
                return false;
            var solid = hidden.ChildElements[0];
            if (solid.LocalName != "solidFill" || solid.NamespaceUri != DrawingNamespace || solid.ChildElements.Count != 1 || solid.GetAttributes().Count != 0)
                return false;
            var color = solid.ChildElements[0];
            if (color.LocalName != "srgbClr" || color.NamespaceUri != DrawingNamespace || color.ChildElements.Count != 0 || color.GetAttributes().Count != 1 ||
                color.GetAttributes()[0].LocalName != "val" || color.GetAttributes()[0].NamespaceUri.Length != 0 || !IsHexColor(color.GetAttributes()[0].Value ?? string.Empty)) return false;
        }
        return true;
    }

    private static bool IsHexColor(string value) => value.Length == 6 && value.All(Uri.IsHexDigit);

    private static bool UIntInRange(string value, uint min, uint max) =>
        uint.TryParse(value, out var parsed) && parsed >= min && parsed <= max;

    private static bool RelationshipNamespace(string value) => PptxNativeObjectCatalog.IsRelationshipNamespace(value);

    private static PresentationNonVisualAccessibility? Accessibility(PresentationImage image)
    {
        if (image.AccessibilityTitle.Length == 0 && image.AltText.Length == 0 && !image.HasAccessibilityDecorative) return null;
        var value = new PresentationNonVisualAccessibility();
        if (image.AccessibilityTitle.Length > 0) value.Title = image.AccessibilityTitle;
        if (image.AltText.Length > 0) value.Description = image.AltText;
        if (image.HasAccessibilityDecorative) value.Decorative = image.AccessibilityDecorative;
        return value;
    }

    private static bool AccessibilityEqual(PresentationImage left, PresentationImage right) =>
        left.AccessibilityTitle.Equals(right.AccessibilityTitle, StringComparison.Ordinal) &&
        left.AltText.Equals(right.AltText, StringComparison.Ordinal) &&
        left.HasAccessibilityDecorative == right.HasAccessibilityDecorative &&
        left.AccessibilityDecorative == right.AccessibilityDecorative;

    private static bool CropSupported(A.SourceRectangle source)
    {
        var known = new HashSet<string>(StringComparer.Ordinal) { "l", "t", "r", "b" };
        return !source.HasChildren && source.GetAttributes().All(attribute => known.Contains(attribute.LocalName)) &&
               CropValuesValid(ReadCrop(source));
    }

    private static PresentationImageCrop ReadCrop(A.SourceRectangle source) => new()
    {
        LeftThousandthPercent = source.Left?.Value ?? 0,
        TopThousandthPercent = source.Top?.Value ?? 0,
        RightThousandthPercent = source.Right?.Value ?? 0,
        BottomThousandthPercent = source.Bottom?.Value ?? 0,
    };

    private static bool CropValuesValid(PresentationImageCrop crop) =>
        crop.LeftThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.TopThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.RightThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.BottomThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.LeftThousandthPercent + crop.RightThousandthPercent < 100_000 &&
        crop.TopThousandthPercent + crop.BottomThousandthPercent < 100_000;

    private static A.SourceRectangle BuildCrop(PresentationImageCrop crop) => new()
    {
        Left = crop.LeftThousandthPercent,
        Top = crop.TopThousandthPercent,
        Right = crop.RightThousandthPercent,
        Bottom = crop.BottomThousandthPercent,
    };

    private static void ApplyCrop(P.BlipFill target, PresentationImageCrop? crop)
    {
        var current = target.GetFirstChild<A.SourceRectangle>();
        if (crop is null)
        {
            current?.Remove();
            return;
        }
        var replacement = BuildCrop(crop);
        if (current is not null)
        {
            current.InsertAfterSelf(replacement);
            current.Remove();
            return;
        }
        var blip = target.GetFirstChild<A.Blip>() ?? throw new CodecException("unsupported_presentation_edit", "Presentation picture lost its embedded blip.");
        blip.InsertAfterSelf(replacement);
    }

    private static bool TransformSupported(A.Transform2D? transform)
    {
        if (transform is null || transform.ChildElements.Count != 2 ||
            transform.Elements<A.Offset>().SingleOrDefault() is not { } offset ||
            transform.Elements<A.Extents>().SingleOrDefault() is not { } extents ||
            offset.X is null || offset.Y is null || extents.Cx is null || extents.Cy is null ||
            offset.HasChildren || extents.HasChildren || offset.GetAttributes().Count != 2 || extents.GetAttributes().Count != 2)
            return false;
        var known = new HashSet<string>(StringComparer.Ordinal) { "rot", "flipH", "flipV" };
        if (transform.GetAttributes().Any(attribute => !known.Contains(attribute.LocalName))) return false;
        var rotation = transform.Rotation?.Value;
        return rotation is null || Math.Abs((long)rotation.Value) <= 21_600_000L;
    }

    private static bool FrameCoordinateSupported(long value, bool allowNegative) =>
        value <= MaxFrameCoordinateEmu && value >= (allowNegative ? -MaxFrameCoordinateEmu : 0L);

    private static bool FrameExtentSupported(long value) => value > 0 && value <= MaxFrameCoordinateEmu;

    private static PresentationImageTransform? ReadTransform(A.Transform2D source)
    {
        var result = new PresentationImageTransform();
        if (source.Rotation is not null) result.RotationAngle60000 = source.Rotation.Value;
        if (source.HorizontalFlip is not null) result.FlipHorizontal = source.HorizontalFlip.Value;
        if (source.VerticalFlip is not null) result.FlipVertical = source.VerticalFlip.Value;
        return result.CalculateSize() == 0 ? null : result;
    }

    private static void ValidateTransform(PresentationImageTransform? transform, string elementId)
    {
        if (transform is null) return;
        if (!transform.HasRotationAngle60000 && !transform.HasFlipHorizontal && !transform.HasFlipVertical)
            throw Invalid(elementId, "transform must define rotation or a flip");
        if (transform.HasRotationAngle60000 && Math.Abs((long)transform.RotationAngle60000) > 21_600_000L)
            throw Invalid(elementId, "rotation must be between -360 and 360 degrees");
    }

    private static void ApplyTransform(A.Transform2D target, PresentationImageTransform? source)
    {
        target.Rotation = source?.HasRotationAngle60000 == true ? source.RotationAngle60000 : null;
        target.HorizontalFlip = source?.HasFlipHorizontal == true ? source.FlipHorizontal : null;
        target.VerticalFlip = source?.HasFlipVertical == true ? source.FlipVertical : null;
    }

    private static CodecException Invalid(string elementId, string message) =>
        new("invalid_presentation_image", $"Presentation image {elementId} {message}.");
}
