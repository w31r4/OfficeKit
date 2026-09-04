using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the small Presentation-only c:chartSpace/c:spPr frame profile.  The
// chart area fill used by older PPJ programs remains compatible, while the
// frame message adds the missing image paint, line and outer-shadow ownership
// without wrapping the chart in an unrelated shape.
internal static class PptxChartFrameCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace OfficeKitNs = "urn:office-kit:chart-frame";
    private const string FrameMarkerUri = "urn:office-kit:chart-frame:v1";
    private static readonly IReadOnlySet<string> ShadowAlignments = new HashSet<string>(StringComparer.Ordinal)
    {
        "tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br",
    };

    internal static bool TryRead(
        XElement chartSpace,
        PptxPartContext? context,
        out PresentationChartFrame? frame,
        out bool editable)
    {
        frame = null;
        editable = true;
        var properties = chartSpace.Element(ChartNs + "spPr");
        if (properties is null) return true;
        if (properties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) return false;
        var children = properties.Elements().ToArray();
        var allowed = new[] { DrawingNs + "noFill", DrawingNs + "solidFill", DrawingNs + "gradFill", DrawingNs + "blipFill", DrawingNs + "ln", DrawingNs + "effectLst" };
        if (children.Any(child => !allowed.Contains(child.Name)) ||
            children.Count(child => child.Name is var name && (name == DrawingNs + "noFill" || name == DrawingNs + "solidFill" || name == DrawingNs + "gradFill" || name == DrawingNs + "blipFill")) > 1 ||
            children.Count(child => child.Name == DrawingNs + "ln") > 1 ||
            children.Count(child => child.Name == DrawingNs + "effectLst") > 1)
            return false;

        SpreadsheetChartSurfaceFill? fill = null;
        var paint = children.FirstOrDefault(child => child.Name is var name &&
            (name == DrawingNs + "noFill" || name == DrawingNs + "solidFill" || name == DrawingNs + "gradFill"));
        if (!XlsxChartSurfaceFillCodec.TryReadPaint(paint, out fill)) return false;
        if (!IsFrameMarked(chartSpace)) fill = null;

        PresentationImagePaint? imageFill = null;
        var imagePaint = children.SingleOrDefault(child => child.Name == DrawingNs + "blipFill");
        if (imagePaint is not null)
        {
            if (context is null || !TryReadImagePaint(imagePaint, context, out imageFill)) return false;
            fill = null;
        }

        SpreadsheetChartLineStyleArtifact? line = null;
        if (!XlsxChartSeriesLineStyleCodec.TryReadLine(properties, out line, allowArrowheads: false)) return false;

        PresentationShadow? shadow = null;
        var effects = children.SingleOrDefault(child => child.Name == DrawingNs + "effectLst");
        if (effects is not null && !TryReadShadow(effects, out shadow)) return false;
        if (fill is null && imageFill is null && line is null && shadow is null) return true;
        frame = new PresentationChartFrame();
        if (fill is not null) frame.Fill = fill;
        if (imageFill is not null) frame.ImageFill = imageFill;
        if (line is not null) frame.Line = line;
        if (shadow is not null) frame.Shadow = shadow;
        return true;
    }

    internal static void Validate(PresentationChartFrame? frame, string subject, PptxAssetCatalog? assets = null)
    {
        if (frame is null) return;
        if (frame.Fill is not null && frame.ImageFill is not null)
            throw new CodecException("invalid_chart_frame", $"{subject} cannot combine a surface fill and an image fill.");
        XlsxChartSurfaceFillCodec.Validate(frame.Fill, subject + " fill");
        if (frame.ImageFill is not null) PptxImagePaintCodec.Validate(frame.ImageFill, subject + " image fill", assets);
        XlsxChartSeriesLineStyleCodec.ValidateLine(frame.Line, "presentation", subject, "frame", subject);
        PptxShadowCodec.Validate(frame.Shadow, subject, "chart frame");
    }

    internal static XElement? Element(
        PresentationChartFrame? frame,
        SpreadsheetChartSurfaceFill? legacyFill,
        string subject,
        PptxPartContext? context = null)
    {
        Validate(frame, subject, context?.Assets);
        var effectiveFill = frame?.Fill ?? legacyFill;
        if (frame is null && effectiveFill is null) return null;
        var output = new XElement(ChartNs + "spPr");
        if (frame?.ImageFill is { } imageFill)
        {
            if (context is null)
                throw new CodecException("unsupported_chart_frame_image", $"{subject} image fill requires a ChartPart relationship context.");
            output.Add(ImagePaintElement(imageFill, context.AddEmbeddedPicture(imageFill.AssetId)));
        }
        else if (effectiveFill is not null) output.Add(XlsxChartSurfaceFillCodec.PaintElement(effectiveFill, subject + " fill"));
        if (frame?.Line is not null) output.Add(XlsxChartSeriesLineStyleCodec.Element(frame.Line));
        if (frame?.Shadow is not null) output.Add(ShadowElement(frame.Shadow));
        return output;
    }

    internal static void Patch(
        XElement chartSpace,
        PresentationChartFrame? frame,
        SpreadsheetChartSurfaceFill? legacyFill,
        string subject,
        PptxPartContext? context = null)
    {
        var existing = chartSpace.Element(ChartNs + "spPr");
        var previousImageRelationshipId = ImageRelationshipId(existing);
        if (existing is not null)
        {
            if (!TryRead(chartSpace, context, out _, out var editable) || !editable)
                throw new CodecException("unsupported_chart_edit", $"{subject} uses an unmodeled chart frame graph.");
        }
        var replacement = Element(frame, legacyFill, subject, context);
        if (replacement is null) existing?.Remove();
        else if (existing is not null) existing.ReplaceWith(replacement);
        else
        {
            var extension = chartSpace.Element(ChartNs + "extLst");
            if (extension is null) chartSpace.Add(replacement);
            else extension.AddBeforeSelf(replacement);
        }
        if (context is not null && previousImageRelationshipId is { Length: > 0 })
            context.RemoveIfUnreferenced(previousImageRelationshipId);
        SetFrameMarker(chartSpace, frame is not null);
    }

    private static bool IsFrameMarked(XElement chartSpace) =>
        chartSpace.Element(ChartNs + "extLst")?.Elements(ChartNs + "ext")
            .Any(extension => (string?)extension.Attribute("uri") == FrameMarkerUri &&
                extension.Element(OfficeKitNs + "frame") is not null) == true;

    private static void SetFrameMarker(XElement chartSpace, bool enabled)
    {
        var extList = chartSpace.Element(ChartNs + "extLst");
        var marker = extList?.Elements(ChartNs + "ext")
            .FirstOrDefault(extension => (string?)extension.Attribute("uri") == FrameMarkerUri);
        if (!enabled)
        {
            marker?.Remove();
            if (extList is not null && !extList.Elements().Any() && !extList.Nodes().Any(node => node is XText text && !string.IsNullOrWhiteSpace(text.Value)))
                extList.Remove();
            return;
        }
        if (marker is not null) return;
        extList ??= new XElement(ChartNs + "extLst");
        marker = new XElement(
            ChartNs + "ext",
            new XAttribute("uri", FrameMarkerUri),
            new XElement(OfficeKitNs + "frame"));
        marker.SetAttributeValue(XNamespace.Xmlns + "officekit", OfficeKitNs);
        extList.Add(marker);
        if (extList.Parent is null)
        {
            var extensionAnchor = chartSpace.Element(ChartNs + "spPr");
            if (extensionAnchor is not null) extensionAnchor.AddAfterSelf(extList);
            else chartSpace.Add(extList);
        }
    }

    private static string? ImageRelationshipId(XElement? properties) =>
        properties?.Element(DrawingNs + "blipFill")?.Element(DrawingNs + "blip")?.Attribute(RelationshipsNs + "embed")?.Value;

    private static bool TryReadImagePaint(
        XElement source,
        PptxPartContext context,
        out PresentationImagePaint paint)
    {
        paint = new PresentationImagePaint();
        if (source.Name != DrawingNs + "blipFill" ||
            source.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration))
            return false;
        var children = source.Elements().ToArray();
        if (children.Length is < 2 or > 3 || children[0].Name != DrawingNs + "blip") return false;
        var blip = children[0];
        var embed = blip.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (embed.Length != 1 || embed[0].Name != RelationshipsNs + "embed" || string.IsNullOrWhiteSpace(embed[0].Value) ||
            blip.Elements().Any(child => child.Name != DrawingNs + "alphaModFix") ||
            blip.Elements(DrawingNs + "alphaModFix").Count() > 1)
            return false;

        uint? opacity = null;
        if (blip.Elements(DrawingNs + "alphaModFix").SingleOrDefault() is { } alpha)
        {
            var attributes = alpha.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
            if (alpha.HasElements || attributes.Length != 1 || attributes[0].Name != "amt" ||
                !uint.TryParse(attributes[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var amount) || amount > 100_000)
                return false;
            opacity = amount;
        }

        PresentationImageCrop? crop = null;
        var modeElement = children[^1];
        if (children.Length == 3)
        {
            if (children[1].Name != DrawingNs + "srcRect" || !TryReadCrop(children[1], out crop)) return false;
        }
        var mode = modeElement.Name == DrawingNs + "tile" &&
                   !modeElement.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) &&
                   !modeElement.HasElements
            ? PresentationImagePaint.Types.Mode.Tile
            : modeElement.Name == DrawingNs + "stretch" &&
              !modeElement.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) &&
              modeElement.Elements().Take(2).Count() == 1 &&
              modeElement.Elements().Single().Name == DrawingNs + "fillRect" &&
              !modeElement.Elements().Single().HasAttributes &&
              !modeElement.Elements().Single().HasElements
                ? PresentationImagePaint.Types.Mode.Stretch
                : PresentationImagePaint.Types.Mode.Unspecified;
        if (mode == PresentationImagePaint.Types.Mode.Unspecified) return false;

        try { paint.AssetId = context.ReadEmbeddedPicture(embed[0].Value).Id; }
        catch (CodecException) { paint = new PresentationImagePaint(); return false; }
        paint.Mode = mode;
        if (crop is not null) paint.Crop = crop;
        if (opacity.HasValue) paint.OpacityThousandthPercent = opacity.Value;
        return true;
    }

    private static bool TryReadCrop(XElement source, out PresentationImageCrop? crop)
    {
        crop = null;
        if (source.HasElements || source.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                (attribute.Name.NamespaceName.Length != 0 || attribute.Name.LocalName is not ("l" or "t" or "r" or "b")))) return false;
        var result = new PresentationImageCrop();
        foreach (var (name, setter) in new (string Name, Action<int> Set)[]
        {
            ("l", value => result.LeftThousandthPercent = value),
            ("t", value => result.TopThousandthPercent = value),
            ("r", value => result.RightThousandthPercent = value),
            ("b", value => result.BottomThousandthPercent = value),
        })
        {
            var raw = (string?)source.Attribute(name);
            if (raw is null) continue;
            if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) || value is < -100_000 or > 100_000)
                return false;
            setter(value);
        }
        if (!PptxImagePaintCodec.CropValuesValid(result)) return false;
        crop = result;
        return true;
    }

    private static XElement ImagePaintElement(PresentationImagePaint paint, string relationshipId)
    {
        var blip = new XElement(DrawingNs + "blip", new XAttribute(RelationshipsNs + "embed", relationshipId));
        if (paint.HasOpacityThousandthPercent)
            blip.Add(new XElement(DrawingNs + "alphaModFix", new XAttribute("amt", paint.OpacityThousandthPercent)));
        var output = new XElement(DrawingNs + "blipFill", blip);
        if (paint.Crop is not null)
        {
            if (!PptxImagePaintCodec.CropValuesValid(paint.Crop))
                throw new CodecException("invalid_chart_frame", "chart frame image crop is outside the bounded signed source rectangle.");
            output.Add(new XElement(DrawingNs + "srcRect",
                new XAttribute("l", paint.Crop.LeftThousandthPercent),
                new XAttribute("t", paint.Crop.TopThousandthPercent),
                new XAttribute("r", paint.Crop.RightThousandthPercent),
                new XAttribute("b", paint.Crop.BottomThousandthPercent)));
        }
        output.Add(paint.Mode == PresentationImagePaint.Types.Mode.Tile
            ? new XElement(DrawingNs + "tile")
            : new XElement(DrawingNs + "stretch", new XElement(DrawingNs + "fillRect")));
        return output;
    }

    private static bool TryReadShadow(XElement effectList, out PresentationShadow? shadow)
    {
        shadow = null;
        if (effectList.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || effectList.Elements().Take(2).Count() != 1)
            return false;
        var outer = effectList.Elements().Single();
        if (outer.Name != DrawingNs + "outerShdw" ||
            outer.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.LocalName is not ("blurRad" or "dist" or "dir" or "algn" or "rotWithShape")) ||
            outer.Elements().Take(2).Count() != 1)
            return false;
        if (!TryLongAttribute(outer, "blurRad", 0, long.MaxValue, out var blur) ||
            !TryLongAttribute(outer, "dist", 0, long.MaxValue, out var distance) ||
            !TryLongAttribute(outer, "dir", 0, 21_599_999, out var direction)) return false;
        var alignment = (string?)outer.Attribute("algn");
        if (alignment is not null && !ShadowAlignments.Contains(alignment)) return false;
        var rotate = (string?)outer.Attribute("rotWithShape");
        var rotateWithShape = false;
        if (rotate is not null && !TryBoolean(rotate, out rotateWithShape)) return false;
        var color = outer.Elements().Single();
        if (color.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") ||
            (color.Name != DrawingNs + "srgbClr" && color.Name != DrawingNs + "schemeClr")) return false;
        var colorValue = (string?)color.Attribute("val");
        string? rgb = null;
        string? scheme = null;
        if (color.Name == DrawingNs + "srgbClr")
        {
            if (colorValue is not { Length: 6 } || !colorValue.All(Uri.IsHexDigit)) return false;
            rgb = colorValue.ToUpperInvariant();
        }
        else
        {
            if (colorValue is null || !PptxColor.TrySchemeToken(colorValue, out var token)) return false;
            scheme = token;
        }
        var alpha = color.Elements(DrawingNs + "alpha").ToArray();
        if (color.Elements().Any(child => child.Name != DrawingNs + "alpha") || alpha.Length > 1) return false;
        uint? opacity = null;
        if (alpha.Length == 1)
        {
            if (alpha[0].HasElements || alpha[0].Attributes().Any(attribute => attribute.Name != "val") ||
                !uint.TryParse((string?)alpha[0].Attribute("val"), NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) || parsed > 100_000) return false;
            opacity = parsed;
        }
        shadow = new PresentationShadow();
        if (rgb is not null) shadow.ColorRgb = rgb;
        else shadow.ColorScheme = scheme!;
        if (blur.HasValue) shadow.BlurRadiusEmu = blur.Value;
        if (distance.HasValue) shadow.DistanceEmu = distance.Value;
        if (direction.HasValue) shadow.DirectionAngle60000 = checked((int)direction.Value);
        if (opacity.HasValue) shadow.OpacityThousandthPercent = opacity.Value;
        if (alignment is not null) shadow.Alignment = alignment;
        if (rotate is not null) shadow.RotateWithShape = rotateWithShape;
        return true;
    }

    private static XElement ShadowElement(PresentationShadow shadow)
    {
        PptxShadowCodec.Validate(shadow, "chart-frame", "chart frame");
        var color = !string.IsNullOrEmpty(shadow.ColorScheme)
            ? new XElement(DrawingNs + "schemeClr", new XAttribute("val", PptxColor.SchemeValue(shadow.ColorScheme)))
            : new XElement(DrawingNs + "srgbClr", new XAttribute("val", PptxColor.Normalize(shadow.ColorRgb)));
        if (shadow.HasOpacityThousandthPercent) color.Add(new XElement(DrawingNs + "alpha", new XAttribute("val", shadow.OpacityThousandthPercent)));
        var outer = new XElement(DrawingNs + "outerShdw", color);
        if (shadow.HasBlurRadiusEmu) outer.SetAttributeValue("blurRad", shadow.BlurRadiusEmu);
        if (shadow.HasDistanceEmu) outer.SetAttributeValue("dist", shadow.DistanceEmu);
        if (shadow.HasDirectionAngle60000) outer.SetAttributeValue("dir", shadow.DirectionAngle60000);
        if (shadow.HasAlignment) outer.SetAttributeValue("algn", shadow.Alignment);
        if (shadow.HasRotateWithShape) outer.SetAttributeValue("rotWithShape", shadow.RotateWithShape ? "1" : "0");
        return new XElement(DrawingNs + "effectLst", outer);
    }

    private static bool TryLongAttribute(XElement owner, string name, long minimum, long maximum, out long? value)
    {
        value = null;
        var raw = (string?)owner.Attribute(name);
        if (raw is null) return true;
        if (!long.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) || parsed < minimum || parsed > maximum) return false;
        value = parsed;
        return true;
    }

    private static bool TryBoolean(string value, out bool result)
    {
        if (value is "1" or "true") { result = true; return true; }
        if (value is "0" or "false") { result = false; return true; }
        result = false;
        return false;
    }
}
