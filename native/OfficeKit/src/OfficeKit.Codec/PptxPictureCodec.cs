using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the deliberately bounded, source-preserving p:pic projection. The
// semantic image owns a compact canonical picture profile: asset, frame, crop,
// accessibility, opacity, preset or bounded custom mask, border, and one outer shadow. SVG
// fallback and bounded presentation extensions remain source-preserved.
internal static class PptxPictureCodec
{
    private const int MaxTextLength = 1_024;
    // A source-bound image may legitimately bleed past the slide edge. Keep
    // that valid DrawingML geometry bounded so malformed imports cannot turn
    // into unbounded JS coordinates or allocations.
    private const long MaxFrameCoordinateEmu = 100_000_000L;
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Office2010Namespace = "http://schemas.microsoft.com/office/drawing/2010/main";
    private const string SvgBlipNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private const string UseLocalDpiUri = "{28A0092B-C50C-407E-A947-70E740481C1C}";
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    private const string HiddenFillUri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}";

    internal static bool TryRead(P.Picture source, PptxPartContext context, out PresentationImage image)
    {
        image = new PresentationImage();
        if (!TryParts(source, out var nonVisual, out var blip, out var properties, out var transform, out var geometry, out var crop, out var tiled) ||
            !TryReadBorder(properties.GetFirstChild<A.Outline>(), out var border) ||
            !PptxShadowCodec.TryRead(properties, out var shadow)) return false;
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
            var svgRelationshipId = SvgFallbackRelationshipId(blip);
            var svgAsset = svgRelationshipId is null ? null : context.ReadEmbeddedPicture(svgRelationshipId);
            if (svgAsset is not null && !svgAsset.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                return false;
            var offset = transform.Offset!;
            var extents = transform.Extents!;
            image = new PresentationImage
            {
                AssetId = asset.Id,
                SvgAssetId = svgAsset?.Id ?? string.Empty,
                AltText = accessibility?.HasDescription == true ? accessibility.Description : string.Empty,
                AccessibilityTitle = accessibility?.HasTitle == true ? accessibility.Title : string.Empty,
                LeftEmu = offset.X?.Value ?? 0,
                TopEmu = offset.Y?.Value ?? 0,
                WidthEmu = extents.Cx?.Value ?? 0,
                HeightEmu = extents.Cy?.Value ?? 0,
            };
            if (accessibility?.HasDecorative == true) image.AccessibilityDecorative = accessibility.Decorative;
            if (crop is not null) image.Crop = ReadCrop(crop);
            image.Tiled = tiled;
            // An empty alphaModFix is a valid no-op effect, not an explicit
            // opacity token. Preserve it as opaque source structure rather
            // than manufacturing an editable 100% leaf that the writer
            // cannot later prove. Only an authored `amt` is a controlled
            // source-bound opacity value.
            if (blip.GetFirstChild<A.AlphaModulationFixed>() is { Amount.Value: { } amount })
                image.OpacityThousandthPercent = checked((uint)amount);
            if (!TryReadMask(geometry, image.WidthEmu, image.HeightEmu, image)) return false;
            image.Border = border;
            image.Shadow = shadow;
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
        if (image.SvgAssetId.Length > 512)
            throw Invalid(elementId, "SVG fallback asset ID must contain at most 512 characters");
        if (image.SvgAssetId.Length > 0 && !assets.Get(image.SvgAssetId).ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
            throw Invalid(elementId, "SVG fallback asset must use content type image/svg+xml");
        PptxNonVisualAccessibilityCodec.Validate(Accessibility(image), elementId, "image");
        if (!FrameCoordinateSupported(image.LeftEmu, sourceBound) || !FrameCoordinateSupported(image.TopEmu, sourceBound) ||
            !FrameExtentSupported(image.WidthEmu) || !FrameExtentSupported(image.HeightEmu))
            throw Invalid(elementId, sourceBound
                ? "source-bound frame coordinates must stay within the bounded DrawingML range and extents must be positive"
                : "frame must use non-negative coordinates and positive extents");
        if (image.Crop is not null && !CropValuesValid(image.Crop))
            throw Invalid(elementId, "crop edges must be between -100% and 100% and opposing sums must remain below 100%");
        if (image.HasOpacityThousandthPercent && image.OpacityThousandthPercent > 100_000)
            throw Invalid(elementId, "opacity must be between 0% and 100%");
        if (image.CustomMaskPaths.Count > 0)
        {
            if (image.MaskPreset.Length > 0 || image.MaskPresetAdjustments.Count > 0)
                throw Invalid(elementId, "custom image mask paths cannot be combined with preset mask state");
            PptxCustomGeometryCodec.Validate(CustomMaskShape(image), elementId + " image mask");
        }
        else
        {
            var maskPreset = image.MaskPreset.Length == 0 ? "rect" : image.MaskPreset;
            PptxPresetGeometryAdjustmentCodec.Validate(maskPreset, image.MaskPresetAdjustments, elementId + " image mask");
        }
        ValidateBorder(image.Border, elementId);
        PptxShadowCodec.Validate(image.Shadow, elementId, "image");
        ValidateTransform(image.Transform, elementId);
    }

    internal static P.Picture Build(PresentationElement source, uint nativeId, PptxPartContext context)
    {
        var image = source.Image;
        // SVG fallbacks are currently source-bound only. Refuse an authored
        // image carrying one instead of silently dropping the fallback while
        // constructing a new picture.
        if (image.SvgAssetId.Length > 0)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author an SVG fallback outside the bounded source-preserving path.");
        var transform = new A.Transform2D(
            new A.Offset { X = image.LeftEmu, Y = image.TopEmu },
            new A.Extents { Cx = image.WidthEmu, Cy = image.HeightEmu });
        ApplyTransform(transform, image.Transform);
        var blip = new A.Blip { Embed = context.AddEmbeddedPicture(image.AssetId) };
        ApplyOpacity(blip, image.HasOpacityThousandthPercent ? image.OpacityThousandthPercent : null);
        var fill = new P.BlipFill(blip);
        if (image.Crop is not null) fill.Append(BuildCrop(image.Crop));
        fill.Append(image.Tiled ? new A.Tile() : new A.Stretch(new A.FillRectangle()));
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, Accessibility(image));
        var properties = new P.ShapeProperties(transform);
        if (image.CustomMaskPaths.Count > 0)
            PptxCustomGeometryCodec.Apply(properties, CustomMaskShape(image), source.Id + " image mask");
        else
            properties.Append(BuildMask(image.MaskPreset, image.MaskPresetAdjustments, source.Id));
        if (image.Border is not null) properties.Append(BuildBorder(image.Border));
        PptxShadowCodec.Apply(properties, image.Shadow);
        return new P.Picture(
            new P.NonVisualPictureProperties(
                nonVisual,
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            fill,
            properties);
    }

    internal static void Apply(P.Picture source, PresentationElement requested, PptxPartContext context)
    {
        if (!TryRead(source, context, out var currentImage) ||
            !TryParts(source, out var nonVisual, out var blip, out var properties, out var transform, out var geometry, out _, out _))
            throw new CodecException("unsupported_presentation_edit", $"Presentation image {requested.Id} no longer matches the editable picture profile.");
        var current = context.ReadEmbeddedPicture(blip.Embed?.Value ?? string.Empty);
        var replacement = context.Assets?.Get(requested.Image.AssetId) ??
            throw new CodecException("invalid_presentation_asset", $"Presentation image {requested.Id} requires an asset catalog.");
        if (!current.Id.Equals(replacement.Id, StringComparison.Ordinal))
        {
            if (!current.ContentType.Equals(replacement.ContentType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} replacement must retain content type {current.ContentType}.");
            if (HasSvgFallback(blip) && requested.Image.SvgAssetId.Length == 0)
                throw new CodecException(
                    "unsupported_presentation_image",
                    $"Presentation image {requested.Id} has a paired SVG fallback; replacing only the primary asset would leave the fallback stale.");
            blip.Embed = context.AddEmbeddedPicture(replacement.Id);
        }
        var currentSvgRelationshipId = SvgFallbackRelationshipId(blip);
        if (currentSvgRelationshipId is null)
        {
            if (requested.Image.SvgAssetId.Length > 0)
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} has no source SVG fallback relationship to replace.");
        }
        else
        {
            var currentSvg = context.ReadEmbeddedPicture(currentSvgRelationshipId);
            var requestedSvg = requested.Image.SvgAssetId.Length == 0
                ? throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} cannot remove a paired SVG fallback in the bounded source-preserving path.")
                : context.Assets.Get(requested.Image.SvgAssetId);
            if (!currentSvg.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase) ||
                !requestedSvg.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} SVG fallback must retain content type image/svg+xml.");
            if (!currentSvg.Id.Equals(requestedSvg.Id, StringComparison.Ordinal))
            {
                var fallback = SvgFallbackElement(blip) ??
                    throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} SVG fallback relationship is malformed.");
                fallback.SetAttribute(new OpenXmlAttribute("r", "embed", OfficeRelationshipsNamespace, context.AddEmbeddedPicture(requestedSvg.Id)));
            }
        }
        nonVisual.Name = requested.Name;
        if (!AccessibilityEqual(currentImage, requested.Image))
            PptxNonVisualAccessibilityCodec.ApplyResidualBound(nonVisual, Accessibility(requested.Image), "image");
        transform.Offset!.X = requested.Image.LeftEmu;
        transform.Offset.Y = requested.Image.TopEmu;
        transform.Extents!.Cx = requested.Image.WidthEmu;
        transform.Extents.Cy = requested.Image.HeightEmu;
        ApplyCrop(source.BlipFill!, requested.Image.Crop);
        if (currentImage.Tiled != requested.Image.Tiled)
        {
            source.BlipFill!.ChildElements.Last().Remove();
            source.BlipFill.Append(requested.Image.Tiled ? new A.Tile() : new A.Stretch(new A.FillRectangle()));
        }
        ApplyTransform(transform, requested.Image.Transform);
        if (currentImage.HasOpacityThousandthPercent != requested.Image.HasOpacityThousandthPercent ||
            currentImage.OpacityThousandthPercent != requested.Image.OpacityThousandthPercent)
            ApplyOpacity(blip, requested.Image.HasOpacityThousandthPercent ? requested.Image.OpacityThousandthPercent : null);
        if (!currentImage.CustomMaskPaths.SequenceEqual(requested.Image.CustomMaskPaths))
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation image {requested.Id} custom mask path topology is source-owned and cannot be changed.");
        if (!currentImage.MaskPreset.Equals(requested.Image.MaskPreset, StringComparison.Ordinal) ||
            !currentImage.MaskPresetAdjustments.SequenceEqual(requested.Image.MaskPresetAdjustments))
        {
            if (geometry is not A.PresetGeometry presetGeometry)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} custom mask identity is source-owned and cannot be replaced by a preset mask.");
            var maskPreset = requested.Image.MaskPreset.Length == 0 ? "rect" : requested.Image.MaskPreset;
            presetGeometry.Preset = MaskPreset(requested.Image.MaskPreset);
            PptxPresetGeometryAdjustmentCodec.Apply(presetGeometry, maskPreset, requested.Image.MaskPresetAdjustments, requested.Id + " image mask");
        }
        if (!Equals(currentImage.Border, requested.Image.Border)) ApplyBorder(properties, requested.Image.Border);
        if (!Equals(currentImage.Shadow, requested.Image.Shadow)) PptxShadowCodec.Apply(properties, requested.Image.Shadow);
    }

    internal static void ScrubModeledContent(P.Picture source)
    {
        if (source.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            nonVisual.Name = string.Empty;
            PptxNonVisualAccessibilityCodec.ScrubResidualModeledContent(nonVisual);
            nonVisual.Description = string.Empty;
        }
        if (source.BlipFill?.GetFirstChild<A.Blip>() is { } blip)
        {
            blip.Embed = string.Empty;
            SvgFallbackElement(blip)?.SetAttribute(new OpenXmlAttribute(
                "r",
                "embed",
                OfficeRelationshipsNamespace,
                string.Empty));
            blip.GetFirstChild<A.AlphaModulationFixed>()?.Remove();
        }
        source.BlipFill?.GetFirstChild<A.SourceRectangle>()?.Remove();
        if (source.BlipFill?.ChildElements.LastOrDefault() is A.Tile or A.Stretch)
        {
            source.BlipFill.ChildElements.Last().Remove();
            source.BlipFill.Append(new A.Stretch(new A.FillRectangle()));
        }
        if (source.ShapeProperties?.Transform2D is { } transform)
        {
            if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
            if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
            transform.Rotation = null;
            transform.HorizontalFlip = null;
            transform.VerticalFlip = null;
        }
        if (source.ShapeProperties is { } properties)
        {
            if (properties.GetFirstChild<A.PresetGeometry>() is { } geometry)
            {
                geometry.Preset = A.ShapeTypeValues.Rectangle;
                geometry.RemoveAllChildren();
                geometry.Append(new A.AdjustValueList());
            }
            properties.GetFirstChild<A.Outline>()?.Remove();
            properties.GetFirstChild<A.EffectList>()?.Remove();
        }
    }

    private static bool TryParts(
        P.Picture source,
        out P.NonVisualDrawingProperties nonVisual,
        out A.Blip blip,
        out P.ShapeProperties properties,
        out A.Transform2D transform,
        out OpenXmlElement geometry,
        out A.SourceRectangle? crop,
        out bool tiled)
    {
        nonVisual = null!;
        blip = null!;
        properties = null!;
        transform = null!;
        geometry = null!;
        crop = null;
        tiled = false;
        var nonVisualContainer = source.NonVisualPictureProperties;
        var fill = source.BlipFill;
        var pictureProperties = source.ShapeProperties;
        if (nonVisualContainer?.NonVisualDrawingProperties is not { } nv ||
            nonVisualContainer.NonVisualPictureDrawingProperties is null ||
            nonVisualContainer.ApplicationNonVisualDrawingProperties is null ||
            fill is null || pictureProperties is null) return false;
        var fillChildren = fill.ChildElements.ToArray();
        if (source.ChildElements.Count != 3 ||
            nonVisualContainer.ChildElements.Count != 3 ||
            fillChildren.Length is < 2 or > 3 || fillChildren[0] is not A.Blip embedded ||
            fillChildren[^1] is not (A.Stretch or A.Tile) ||
            fillChildren.Length == 3 && fillChildren[1] is not A.SourceRectangle ||
            fillChildren[^1] is A.Stretch stretch && !StretchSupported(stretch) ||
            fillChildren[^1] is A.Tile tile && !TileSupported(tile) ||
            !BlipSupported(embedded) || !BlipFillSupported(fill) ||
            !ShapePropertiesSupported(pictureProperties) ||
            pictureProperties.Elements<A.Transform2D>().SingleOrDefault() is not { } xfrm ||
            !TransformSupported(xfrm)) return false;
        var geometries = pictureProperties.ChildElements
            .Where(child => child is A.PresetGeometry or A.CustomGeometry)
            .ToArray();
        if (geometries.Length != 1) return false;
        if (geometries[0] is A.PresetGeometry presetGeometry)
        {
            if (presetGeometry.Preset is null || presetGeometry.GetAttributes().Count != 1 ||
                presetGeometry.GetAttributes()[0].LocalName != "prst" || presetGeometry.GetAttributes()[0].NamespaceUri.Length != 0 ||
                presetGeometry.ChildElements.Count != 1 || presetGeometry.GetFirstChild<A.AdjustValueList>() is not { } adjustments ||
                adjustments.HasAttributes) return false;
        }
        else if (geometries[0] is not A.CustomGeometry customGeometry ||
                 !PptxCustomGeometryCodec.Supports(
                     customGeometry,
                     xfrm.Extents?.Cx?.Value ?? 0,
                     xfrm.Extents?.Cy?.Value ?? 0)) return false;
        crop = fillChildren.Length == 3 ? (A.SourceRectangle)fillChildren[1] : null;
        tiled = fillChildren[^1] is A.Tile;
        if (crop is not null && !CropSupported(crop)) return false;
        nonVisual = nv;
        blip = embedded;
        properties = pictureProperties;
        transform = xfrm;
        geometry = geometries[0];
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

    private static OpenXmlElement? SvgFallbackElement(A.Blip blip) => blip.ChildElements
        .Where(child => child.LocalName == "extLst" && child.NamespaceUri == DrawingNamespace)
        .SelectMany(child => child.ChildElements)
        .Where(child => child.LocalName == "ext" && child.NamespaceUri == DrawingNamespace &&
            child.GetAttributes().Any(attribute => attribute.LocalName == "uri" && attribute.NamespaceUri.Length == 0 && attribute.Value == SvgExtensionUri))
        .SelectMany(child => child.ChildElements)
        .SingleOrDefault(child => child.LocalName == "svgBlip" && child.NamespaceUri == SvgBlipNamespace);

    private static string? SvgFallbackRelationshipId(A.Blip blip)
    {
        var fallback = SvgFallbackElement(blip);
        if (fallback is null) return null;
        var embeds = fallback.GetAttributes()
            .Where(attribute => attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri))
            .ToArray();
        return embeds.Length == 1 ? embeds[0].Value : string.Empty;
    }

    private static bool HasSvgFallback(A.Blip blip) => SvgFallbackElement(blip) is not null;

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

    private static bool TileSupported(A.Tile tile) => !tile.HasAttributes && !tile.HasChildren;

    private static bool ShapePropertiesSupported(P.ShapeProperties properties)
    {
        var attributes = properties.GetAttributes().ToArray();
        if (attributes.Any(attribute => attribute.LocalName != "bwMode" || attribute.NamespaceUri.Length != 0 || attribute.Value is not ("auto" or "gray" or "ltGray" or "invGray" or "grayWhite" or "blackGray" or "blackWhite" or "clr"))) return false;
        if (properties.ChildElements.Count < 2 || properties.ChildElements.Count > 6) return false;
        if (properties.ChildElements.Count(child => child is A.Transform2D) != 1 ||
            properties.ChildElements.Count(child => child is A.PresetGeometry or A.CustomGeometry) != 1) return false;
        return properties.ChildElements.All(child => child switch
        {
            A.Transform2D or A.PresetGeometry or A.CustomGeometry => true,
            A.NoFill noFill => !noFill.HasAttributes && !noFill.HasChildren,
            A.Outline outline => TryReadBorder(outline, out _),
            A.EffectList => PptxShadowCodec.TryRead(properties, out _),
            A.ShapePropertiesExtensionList extensions => ShapeExtensionListSupported(extensions),
            _ => false,
        });
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

    private static bool TryReadBorder(A.Outline? outline, out PresentationImageBorder? border)
    {
        border = null;
        if (outline is null) return true;
        // PowerPoint and Google Slides commonly serialize a picture frame's
        // default single/centered line attributes even when the frame is
        // explicitly invisible.  These attributes do not change the rendered
        // picture, so accept only the canonical defaults while retaining the
        // existing source XML until a caller explicitly edits the border.
        if (outline.CompoundLineType?.Value is { } compound &&
            !compound.Equals(A.CompoundLineValues.Single) ||
            outline.Alignment?.Value is { } alignment &&
            !alignment.Equals(A.PenAlignmentValues.Center) ||
            !HasOnlyAttributes(outline, "w", "cap", "cmpd", "algn") ||
            outline.Width?.Value is < 0 or > int.MaxValue)
            return false;
        var noFill = outline.Elements<A.NoFill>().ToArray();
        var solidFill = outline.Elements<A.SolidFill>().ToArray();
        if (noFill.Length + solidFill.Length != 1) return false;
        var noLine = noFill.Length == 1;
        if (noLine && (noFill[0].HasAttributes || noFill[0].HasChildren)) return false;

        string rgb = string.Empty;
        uint? opacity = null;
        if (!noLine)
        {
            var solid = solidFill[0];
            if (solid.ChildElements.Count != 1 || solid.FirstChild is not A.RgbColorModelHex color ||
                !HasOnlyAttributes(solid) || !HasOnlyAttributes(color, "val") ||
                color.Val?.Value is not { Length: 6 } directRgb || !directRgb.All(Uri.IsHexDigit)) return false;
            var alphas = color.Elements<A.Alpha>().ToArray();
            if (color.ChildElements.Count != alphas.Length || alphas.Length > 1 ||
                alphas.Length == 1 && (alphas[0].Val?.Value is not (>= 0 and <= 100_000) || !HasOnlyAttributes(alphas[0], "val")))
                return false;
            rgb = directRgb;
            if (alphas.SingleOrDefault()?.Val?.Value is { } alphaValue)
                opacity = checked((uint)alphaValue);
        }
        var dashes = outline.Elements<A.PresetDash>().ToArray();
        if (dashes.Length > 1 || dashes.SingleOrDefault() is { } dash &&
            (dash.ChildElements.Any() || !HasOnlyAttributes(dash, "val"))) return false;
        if (noLine && dashes.Length > 0) return false;
        var style = dashes.SingleOrDefault()?.Val?.Value switch
        {
            null => "solid",
            var value when value.Equals(A.PresetLineDashValues.Solid) => "solid",
            var value when value.Equals(A.PresetLineDashValues.Dash) => "dashed",
            var value when value.Equals(A.PresetLineDashValues.Dot) => "dotted",
            var value when value.Equals(A.PresetLineDashValues.DashDot) => "dash-dot",
            var value when value.Equals(A.PresetLineDashValues.LargeDashDotDot) => "dash-dot-dot",
            _ => string.Empty,
        };
        if (noLine) style = "none";
        if (style.Length == 0 || !TryCap(outline.CapType?.Value, out var cap) || !TryJoin(outline, out var join) ||
            !TryReadInertLineEnds(outline) ||
            outline.ChildElements.Any(child => child is not A.NoFill and not A.SolidFill and not A.PresetDash and not A.Round and not A.LineJoinBevel and not A.Miter and not A.HeadEnd and not A.TailEnd))
            return false;
        if (noLine) return true;
        border = new PresentationImageBorder
        {
            ColorRgb = PptxColor.Normalize(rgb),
            WidthEmu = outline.Width?.Value ?? 0L,
            Style = style,
            Cap = cap,
            Join = join,
        };
        if (opacity is { } alpha)
            border.OpacityThousandthPercent = alpha;
        return true;
    }

    private static void ValidateBorder(PresentationImageBorder? border, string elementId)
    {
        if (border is null) return;
        PptxColor.Normalize(border.ColorRgb);
        if (border.WidthEmu is < 0 or > int.MaxValue ||
            border.Style is not ("solid" or "dashed" or "dotted" or "dash-dot" or "dash-dot-dot") ||
            border.Cap is not ("" or "flat" or "round" or "square") ||
            border.Join is not ("" or "miter" or "round" or "bevel") ||
            border.HasOpacityThousandthPercent && border.OpacityThousandthPercent > 100_000)
            throw Invalid(elementId, "border uses unsupported width, dash, cap, join, or opacity");
    }

    private static A.Outline BuildBorder(PresentationImageBorder border)
    {
        var outline = new A.Outline { Width = checked((int)border.WidthEmu) };
        outline.CapType = border.Cap switch
        {
            "round" => A.LineCapValues.Round,
            "square" => A.LineCapValues.Square,
            "flat" => A.LineCapValues.Flat,
            _ => null,
        };
        var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(border.ColorRgb) };
        if (border.HasOpacityThousandthPercent)
            color.Append(new A.Alpha { Val = checked((int)border.OpacityThousandthPercent) });
        outline.Append(new A.SolidFill(color));
        if (border.Style != "solid") outline.Append(new A.PresetDash { Val = border.Style switch
        {
            "dotted" => A.PresetLineDashValues.Dot,
            "dash-dot" => A.PresetLineDashValues.DashDot,
            "dash-dot-dot" => A.PresetLineDashValues.LargeDashDotDot,
            _ => A.PresetLineDashValues.Dash,
        }});
        if (border.Join.Length > 0) outline.Append(border.Join switch
        {
            "round" => new A.Round(),
            "bevel" => new A.LineJoinBevel(),
            _ => new A.Miter(),
        });
        return outline;
    }

    private static void ApplyBorder(P.ShapeProperties properties, PresentationImageBorder? border)
    {
        properties.GetFirstChild<A.Outline>()?.Remove();
        if (border is null) return;
        var outline = BuildBorder(border);
        OpenXmlElement? anchor = properties.ChildElements.LastOrDefault(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill);
        anchor ??= properties.GetFirstChild<A.PresetGeometry>();
        anchor ??= properties.GetFirstChild<A.Transform2D>();
        if (anchor is null) properties.PrependChild(outline);
        else properties.InsertAfter(outline, anchor);
    }

    private static A.PresetGeometry BuildMask(string value, IEnumerable<int> adjustments, string elementId)
    {
        var name = value.Length == 0 ? "rect" : value;
        var output = new A.PresetGeometry { Preset = MaskPreset(value) };
        PptxPresetGeometryAdjustmentCodec.Apply(output, name, adjustments, elementId + " image mask");
        return output;
    }

    private static bool TryReadMask(
        OpenXmlElement geometry,
        long widthEmu,
        long heightEmu,
        PresentationImage image)
    {
        if (geometry is A.PresetGeometry presetGeometry)
        {
            if (presetGeometry.Preset?.Value is not { } preset ||
                !PptxCustomGeometryCodec.TryPresetName(preset, out var maskPreset) ||
                !PptxPresetGeometryAdjustmentCodec.TryRead(presetGeometry, maskPreset, out var maskAdjustments))
                return false;
            if (!maskPreset.Equals("rect", StringComparison.Ordinal)) image.MaskPreset = maskPreset;
            image.MaskPresetAdjustments.Add(maskAdjustments);
            return true;
        }
        if (geometry is not A.CustomGeometry customGeometry) return false;
        var shape = new PresentationShape
        {
            Geometry = "custom",
            WidthEmu = widthEmu,
            HeightEmu = heightEmu,
        };
        PptxCustomGeometryCodec.Read(customGeometry, widthEmu, heightEmu, shape);
        if (shape.CustomPaths.Count == 0 || shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 ||
            shape.CustomConnectionSites.Count > 0 || shape.CustomAdjustmentHandles.Count > 0 || shape.TextRectangle is not null)
            return false;
        image.CustomMaskPaths.Add(shape.CustomPaths);
        return true;
    }

    private static PresentationShape CustomMaskShape(PresentationImage image)
    {
        var shape = new PresentationShape
        {
            Geometry = "custom",
            WidthEmu = image.WidthEmu,
            HeightEmu = image.HeightEmu,
        };
        shape.CustomPaths.Add(image.CustomMaskPaths);
        return shape;
    }

    private static A.ShapeTypeValues MaskPreset(string value)
    {
        var name = value.Length == 0 ? "rect" : value;
        if (!PptxCustomGeometryCodec.TryPreset(name, out var preset))
            throw new CodecException("unsupported_presentation_image", $"Presentation image mask preset {name} is unsupported.");
        return preset;
    }

    private static void ApplyOpacity(A.Blip blip, uint? opacity)
    {
        blip.GetFirstChild<A.AlphaModulationFixed>()?.Remove();
        if (opacity is null) return;
        var alpha = new A.AlphaModulationFixed { Amount = checked((int)opacity.Value) };
        var before = blip.ChildElements.FirstOrDefault(child => child is A.ColorChange or A.BlipExtensionList);
        if (before is null) blip.Append(alpha);
        else blip.InsertBefore(alpha, before);
    }

    private static bool TryCap(A.LineCapValues? value, out string cap)
    {
        cap = string.Empty;
        if (value is null) return true;
        if (value.Value.Equals(A.LineCapValues.Flat)) cap = "flat";
        else if (value.Value.Equals(A.LineCapValues.Round)) cap = "round";
        else if (value.Value.Equals(A.LineCapValues.Square)) cap = "square";
        else return false;
        return true;
    }

    private static bool TryJoin(A.Outline outline, out string join)
    {
        join = string.Empty;
        var joins = outline.ChildElements.Where(child => child is A.Round or A.LineJoinBevel or A.Miter).ToArray();
        if (joins.Length > 1) return false;
        if (joins.SingleOrDefault() is A.Round round)
        {
            if (round.HasAttributes || round.HasChildren) return false;
            join = "round";
        }
        else if (joins.SingleOrDefault() is A.LineJoinBevel bevel)
        {
            if (bevel.HasAttributes || bevel.HasChildren) return false;
            join = "bevel";
        }
        else if (joins.SingleOrDefault() is A.Miter miter)
        {
            // 800000 is the DrawingML default miter limit.  It is inert for
            // the picture frame profile, but is emitted explicitly by several
            // native exporters alongside noFill.
            if (miter.HasChildren || miter.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName != "lim" || attribute.Value != "800000")) return false;
            join = "miter";
        }
        return true;
    }

    private static bool TryReadInertLineEnds(A.Outline outline)
    {
        var heads = outline.Elements<A.HeadEnd>().ToArray();
        var tails = outline.Elements<A.TailEnd>().ToArray();
        if (heads.Length > 1 || tails.Length > 1) return false;
        return InertLineEnd(heads.SingleOrDefault()) && InertLineEnd(tails.SingleOrDefault());
    }

    private static bool InertLineEnd(OpenXmlElement? source)
    {
        if (source is null) return true;
        if (source.ChildElements.Any() || !HasOnlyAttributes(source, "type", "w", "len")) return false;
        var type = source switch
        {
            A.HeadEnd head => head.Type?.Value,
            A.TailEnd tail => tail.Type?.Value,
            _ => null,
        };
        return type is null || type.Value.Equals(A.LineEndValues.None);
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }

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
