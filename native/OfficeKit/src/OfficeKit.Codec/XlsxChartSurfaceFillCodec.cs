using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded chart-area and plot-area paint profile shared by PPTX and
// XLSX ChartSpace: direct no-fill, direct sRGB solid fill with optional alpha,
// or the same bounded literal gradient used by Presentation shapes.
internal static class XlsxChartSurfaceFillCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartSurfaceFill? fill, string subject)
    {
        if (fill is null) return;
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.None)
            throw new CodecException("invalid_chart_fill", $"{subject} must select no_fill, solid_rgb, or gradient_fill.");
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill && !fill.NoFill)
            throw new CodecException("invalid_chart_fill", $"{subject} no_fill must be true.");
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb)
            _ = PptxColor.Normalize(fill.SolidRgb);
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.GradientFill)
            PptxGradientFillCodec.Validate(fill.GradientFill, subject);
        if (fill.HasOpacityThousandthPercent &&
            (fill.FillCase != SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb || fill.OpacityThousandthPercent > 100_000))
            throw new CodecException("invalid_chart_fill", $"{subject} opacity requires a solid fill from 0 to 100000.");
    }

    internal static bool TryRead(XElement? properties, out SpreadsheetChartSurfaceFill? fill) =>
        TryRead(properties, out fill, allowFrameDecorations: false);

    // Presentation charts own line/effect children in the separate
    // chartFrame profile.  Let that caller inspect the fill without marking
    // the entire chart read-only, while the XLSX/shared route keeps the
    // historical fill-only contract.
    internal static bool TryRead(
        XElement? properties,
        out SpreadsheetChartSurfaceFill? fill,
        bool allowFrameDecorations)
    {
        fill = null;
        if (properties is null) return true;
        if (properties.HasAttributes) return false;
        var children = properties.Elements().ToArray();
        if (!allowFrameDecorations)
        {
            if (children.Length != 1) return false;
            return TryReadPaint(children[0], out fill);
        }
        if (children.Any(child => child.Name != DrawingNs + "noFill" &&
                                  child.Name != DrawingNs + "solidFill" &&
                                  child.Name != DrawingNs + "gradFill" &&
                                  child.Name != DrawingNs + "ln" &&
                                  child.Name != DrawingNs + "effectLst") ||
            children.Count(child => child.Name is var name &&
                                   (name == DrawingNs + "noFill" || name == DrawingNs + "solidFill" || name == DrawingNs + "gradFill")) > 1)
            return false;
        var paint = children.FirstOrDefault(child => child.Name == DrawingNs + "noFill" ||
                                                     child.Name == DrawingNs + "solidFill" ||
                                                     child.Name == DrawingNs + "gradFill");
        return TryReadPaint(paint, out fill);
    }

    internal static bool TryReadPaint(XElement? paint, out SpreadsheetChartSurfaceFill? fill)
    {
        fill = null;
        if (paint is null) return true;
        if (paint.Name == DrawingNs + "noFill" && !paint.HasAttributes && !paint.HasElements)
        {
            fill = new SpreadsheetChartSurfaceFill { NoFill = true };
            return true;
        }
        if (paint.Name == DrawingNs + "gradFill")
        {
            if (!TryReadGradient(paint, out var gradient)) return false;
            fill = new SpreadsheetChartSurfaceFill { GradientFill = gradient };
            return true;
        }
        if (paint.Name != DrawingNs + "solidFill" || paint.HasAttributes || paint.Elements().Take(2).Count() != 1)
            return false;
        var color = paint.Elements().Single();
        var rgb = (string?)color.Attribute("val");
        if (color.Name != DrawingNs + "srgbClr" || rgb is not { Length: 6 } || !rgb.All(Uri.IsHexDigit) ||
            color.Attributes().Count() != 1)
            return false;
        var alpha = color.Elements(DrawingNs + "alpha").Take(2).ToArray();
        if (color.Elements().Any(child => child.Name != DrawingNs + "alpha") || alpha.Length > 1)
            return false;
        fill = new SpreadsheetChartSurfaceFill { SolidRgb = rgb.ToUpperInvariant() };
        if (alpha.Length == 1)
        {
            var value = (string?)alpha[0].Attribute("val");
            if (alpha[0].HasElements || alpha[0].Attributes().Count() != 1 ||
                !uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var opacity) || opacity > 100_000)
                return false;
            fill.OpacityThousandthPercent = opacity;
        }
        return true;
    }

    internal static XElement? Element(SpreadsheetChartSurfaceFill? fill, string subject)
    {
        if (fill is null) return null;
        Validate(fill, subject);
        return new XElement(ChartNs + "spPr", PaintElement(fill, subject));
    }

    internal static XElement PaintElement(SpreadsheetChartSurfaceFill fill, string subject)
    {
        Validate(fill, subject);
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill)
            return new XElement(DrawingNs + "noFill");
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.GradientFill)
            return GradientElement(fill.GradientFill);
        var color = new XElement(DrawingNs + "srgbClr", new XAttribute("val", PptxColor.Normalize(fill.SolidRgb)));
        if (fill.HasOpacityThousandthPercent)
            color.Add(new XElement(DrawingNs + "alpha", new XAttribute("val", fill.OpacityThousandthPercent)));
        return new XElement(DrawingNs + "solidFill", color);
    }

    internal static void Patch(XElement owner, SpreadsheetChartSurfaceFill? fill, string subject) =>
        Patch(owner, fill, subject, allowFrameDecorations: false);

    internal static void Patch(
        XElement owner,
        SpreadsheetChartSurfaceFill? fill,
        string subject,
        bool allowFrameDecorations)
    {
        var existing = owner.Element(ChartNs + "spPr");
        SpreadsheetChartSurfaceFill? current = null;
        if (existing is not null && !TryRead(existing, out current, allowFrameDecorations))
        {
            // A Presentation ChartPart can carry an outer a:blipFill. Its
            // relationship-aware semantics are patched by PptxChartFrameCodec;
            // the shared chart-area surface must leave it untouched when the
            // requested legacy fill is absent, and must fail closed if callers
            // try to reinterpret that same paint as a worksheet-style fill.
            if (allowFrameDecorations && existing.Elements(DrawingNs + "blipFill").Any())
            {
                if (fill is null) return;
                throw new CodecException("unsupported_chart_edit", $"{subject} cannot replace a ChartPart image frame with a chart-area surface fill.");
            }
            throw new CodecException("unsupported_chart_edit", $"{subject} uses an unmodeled fill graph.");
        }
        if (existing is not null && Semantics(current) == Semantics(fill)) return;
        if (allowFrameDecorations && existing is not null)
        {
            var currentPaint = existing.Elements().FirstOrDefault(IsPaint);
            if (fill is null)
            {
                currentPaint?.Remove();
                if (!existing.Elements().Any()) existing.Remove();
                return;
            }
            var replacementPaint = PaintElement(fill, subject);
            if (currentPaint is not null) currentPaint.ReplaceWith(replacementPaint);
            else existing.AddFirst(replacementPaint);
            return;
        }
        var replacement = Element(fill, subject);
        if (replacement is null)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null) existing.ReplaceWith(replacement);
        else
        {
            var extension = owner.Element(ChartNs + "extLst");
            if (extension is null) owner.Add(replacement);
            else extension.AddBeforeSelf(replacement);
        }
    }

    internal static string Semantics(SpreadsheetChartSurfaceFill? fill) => fill?.FillCase switch
    {
        SpreadsheetChartSurfaceFill.FillOneofCase.NoFill => "none",
        SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb => string.Join(':',
            PptxColor.Normalize(fill.SolidRgb),
            fill.HasOpacityThousandthPercent ? fill.OpacityThousandthPercent.ToString(CultureInfo.InvariantCulture) : "opaque"),
        SpreadsheetChartSurfaceFill.FillOneofCase.GradientFill => "gradient:" + GradientSemantics(fill.GradientFill),
        _ => "default",
    };

    private static bool IsPaint(XElement element) => element.Name == DrawingNs + "noFill" ||
        element.Name == DrawingNs + "solidFill" || element.Name == DrawingNs + "gradFill";

    private static XElement GradientElement(PresentationGradientFill gradient)
    {
        var stops = new XElement(DrawingNs + "gsLst");
        foreach (var stop in gradient.Stops)
        {
            var color = new XElement(DrawingNs + "srgbClr", new XAttribute("val", PptxColor.Normalize(stop.ColorRgb)));
            if (stop.HasOpacityThousandthPercent)
                color.Add(new XElement(DrawingNs + "alpha", new XAttribute("val", stop.OpacityThousandthPercent)));
            stops.Add(new XElement(DrawingNs + "gs", new XAttribute("pos", stop.PositionThousandthPercent), color));
        }
        var geometry = gradient.Kind == PresentationGradientFill.Types.Kind.Linear
            ? new XElement(DrawingNs + "lin",
                new XAttribute("ang", gradient.HasAngle60000 ? gradient.Angle60000 : 0),
                new XAttribute("scaled", "0"))
            : new XElement(DrawingNs + "path",
                new XAttribute("path", "circle"),
                new XElement(DrawingNs + "fillToRect",
                    new XAttribute("l", 50_000),
                    new XAttribute("t", 50_000),
                    new XAttribute("r", 50_000),
                    new XAttribute("b", 50_000)));
        return new XElement(DrawingNs + "gradFill", stops, geometry);
    }

    private static bool TryReadGradient(XElement source, out PresentationGradientFill gradient)
    {
        gradient = new PresentationGradientFill();
        var children = source.Elements().ToArray();
        if (source.HasAttributes || children.Length is < 2 or > 3 || children[0].Name != DrawingNs + "gsLst" ||
            children[0].HasAttributes || children[0].Elements().Count() is < 2 or > 16)
            return false;
        if (children.Length == 3 &&
            (children[2].Name != DrawingNs + "tileRect" || children[2].HasAttributes || children[2].HasElements))
            return false;

        uint previous = 0;
        foreach (var nativeStop in children[0].Elements())
        {
            if (nativeStop.Name != DrawingNs + "gs" || nativeStop.Attributes().Any(attribute => attribute.Name != "pos") ||
                !uint.TryParse((string?)nativeStop.Attribute("pos"), NumberStyles.None, CultureInfo.InvariantCulture, out var position) ||
                position > 100_000 || gradient.Stops.Count > 0 && position < previous ||
                nativeStop.Elements().Take(2).Count() != 1)
                return false;
            var nativeColor = nativeStop.Elements().Single();
            var rgb = (string?)nativeColor.Attribute("val");
            if (nativeColor.Name != DrawingNs + "srgbClr" || rgb is not { Length: 6 } || !rgb.All(Uri.IsHexDigit) ||
                nativeColor.Attributes().Any(attribute => attribute.Name != "val"))
                return false;
            var alphas = nativeColor.Elements(DrawingNs + "alpha").ToArray();
            if (nativeColor.Elements().Any(child => child.Name != DrawingNs + "alpha") || alphas.Length > 1)
                return false;
            var stop = new PresentationGradientStop
            {
                PositionThousandthPercent = position,
                ColorRgb = rgb.ToUpperInvariant(),
            };
            if (alphas.SingleOrDefault() is { } alpha)
            {
                if (alpha.HasElements || alpha.Attributes().Any(attribute => attribute.Name != "val") ||
                    !uint.TryParse((string?)alpha.Attribute("val"), NumberStyles.None, CultureInfo.InvariantCulture, out var opacity) ||
                    opacity > 100_000)
                    return false;
                stop.OpacityThousandthPercent = opacity;
            }
            gradient.Stops.Add(stop);
            previous = position;
        }

        if (children[1].Name == DrawingNs + "lin" && !children[1].HasElements &&
            children[1].Attributes().All(attribute => attribute.Name.LocalName is "ang" or "scaled") &&
            int.TryParse((string?)children[1].Attribute("ang"), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var angle) &&
            angle is >= 0 and < 21_600_000 && (string?)children[1].Attribute("scaled") is "0" or "false")
        {
            gradient.Kind = PresentationGradientFill.Types.Kind.Linear;
            gradient.Angle60000 = angle;
            return true;
        }
        var rectangle = children[1].Element(DrawingNs + "fillToRect");
        if (children[1].Name == DrawingNs + "path" && (string?)children[1].Attribute("path") == "circle" &&
            children[1].Attributes().All(attribute => attribute.Name == "path") &&
            children[1].Elements().Count() == 1 && rectangle is not null && !rectangle.HasElements &&
            rectangle.Attributes().Count() == 4 &&
            new[] { "l", "t", "r", "b" }.All(name => (string?)rectangle.Attribute(name) == "50000"))
        {
            gradient.Kind = PresentationGradientFill.Types.Kind.Radial;
            return true;
        }
        gradient = new PresentationGradientFill();
        return false;
    }

    private static string GradientSemantics(PresentationGradientFill gradient) => string.Join(':',
        gradient.Kind.ToString(),
        gradient.HasAngle60000 ? gradient.Angle60000.ToString(CultureInfo.InvariantCulture) : "no-angle",
        string.Join(',', gradient.Stops.Select(stop => string.Join('@',
            stop.PositionThousandthPercent.ToString(CultureInfo.InvariantCulture),
            PptxColor.Normalize(stop.ColorRgb),
            stop.HasOpacityThousandthPercent ? stop.OpacityThousandthPercent.ToString(CultureInfo.InvariantCulture) : "opaque"))));
}
