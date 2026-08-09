using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the bounded direct a:ln profile of an ordinary p:sp. A free-positioned
// line is still a shape: its p:sp/a:xfrm carries the endpoints while this
// module owns only RGB/no-fill, width, and a small preset-dash vocabulary.
// Arrowheads, joins, caps, compound lines, theme colors, and custom dash graphs
// remain outside this profile so an imported shape containing them is preserved
// as source-bound content and fails closed on mutation.
internal static class PptxShapeLineCodec
{
    internal sealed record Profile(string Rgb, long WidthEmu, string Style);

    private static readonly IReadOnlySet<string> Styles = new HashSet<string>(StringComparer.Ordinal)
    {
        "solid",
        "dashed",
        "dotted",
        "dash-dot",
        "dash-dot-dot",
        "none",
    };

    internal static bool TryRead(A.Outline? outline, out Profile profile)
    {
        if (outline is null)
        {
            profile = new Profile(string.Empty, 0, "none");
            return true;
        }

        profile = null!;
        var width = outline.Width?.Value ?? 0;
        if (width < 0 || width > int.MaxValue || outline.CapType is not null ||
            outline.CompoundLineType is not null || outline.Alignment is not null ||
            !HasOnlyAttributes(outline, "w"))
            return false;

        var noFills = outline.Elements<A.NoFill>().ToArray();
        var solidFills = outline.Elements<A.SolidFill>().ToArray();
        // ISO/IEC 29500 does not define the rendered result when an existing
        // a:ln omits EG_LineFillProperties. Only an explicit noFill or one RGB
        // solidFill is safe to project as editable.
        if (noFills.Length + solidFills.Length != 1) return false;

        var rgb = string.Empty;
        var noFill = noFills.Length == 1;
        if (noFills.Length == 1 && (noFills[0].ChildElements.Any() || !HasOnlyAttributes(noFills[0]))) return false;
        if (solidFills.Length == 1)
        {
            var solid = solidFills[0];
            if (solid.ChildElements.Count != 1 || solid.FirstChild is not A.RgbColorModelHex color ||
                !HasOnlyAttributes(solid) || color.ChildElements.Any() || !HasOnlyAttributes(color, "val") ||
                color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit))
                return false;
            rgb = value.ToUpperInvariant();
        }

        var dashes = outline.Elements<A.PresetDash>().ToArray();
        if (dashes.Length > 1 || dashes.SingleOrDefault() is { } dashElement &&
            (dashElement.ChildElements.Any() || !HasOnlyAttributes(dashElement, "val")))
            return false;
        if (outline.ChildElements.Any(child => child is not A.NoFill and not A.SolidFill and not A.PresetDash))
            return false;

        var dash = dashes.SingleOrDefault()?.Val?.Value;
        if (noFill)
        {
            if (dash is not null) return false;
            profile = new Profile(string.Empty, width, "none");
            return true;
        }

        var style = dash is null || dash.Value.Equals(A.PresetLineDashValues.Solid) ? "solid" :
            dash.Value.Equals(A.PresetLineDashValues.Dash) ? "dashed" :
            dash.Value.Equals(A.PresetLineDashValues.Dot) ? "dotted" :
            dash.Value.Equals(A.PresetLineDashValues.DashDot) ? "dash-dot" :
            dash.Value.Equals(A.PresetLineDashValues.LargeDashDotDot) ? "dash-dot-dot" : string.Empty;
        if (style.Length == 0) return false;
        profile = new Profile(rgb, width, style);
        return true;
    }

    internal static void ReadForProjection(A.Outline? outline, PresentationShape target)
    {
        if (TryRead(outline, out var profile))
        {
            target.LineRgb = profile.Rgb;
            target.LineWidthEmu = profile.WidthEmu;
            target.LineStyle = profile.Style;
            return;
        }

        // The surrounding element remains source-bound. Retain only the old
        // compatibility projection for inspection; it never grants editability.
        target.LineRgb = FallbackRgb(outline?.GetFirstChild<A.SolidFill>());
        var width = outline?.Width?.Value ?? 0;
        target.LineWidthEmu = width is >= 0 and <= int.MaxValue ? width : 0;
        target.LineStyle = target.LineRgb.Length > 0 ? "solid" : "none";
    }

    internal static void Validate(PresentationShape source, string shapeId)
    {
        if (source.LineWidthEmu < 0 || source.LineWidthEmu > int.MaxValue)
            throw new CodecException("invalid_presentation_line", $"Presentation shape {shapeId} has an invalid line width.");
        var style = NormalizeStyle(source.LineStyle, source.LineRgb);
        if (!Styles.Contains(style))
            throw new CodecException("unsupported_presentation_line", $"Presentation shape {shapeId} uses unsupported line style {source.LineStyle}.");
        if (style == "none")
        {
            if (!string.IsNullOrWhiteSpace(source.LineRgb))
                throw new CodecException("invalid_presentation_line", $"Presentation shape {shapeId} cannot combine line style none with a color.");
            return;
        }
        if (string.IsNullOrWhiteSpace(source.LineRgb))
            throw new CodecException("invalid_presentation_line", $"Presentation shape {shapeId} line style {style} requires an RGB color.");
        _ = PptxColor.Normalize(source.LineRgb);
    }

    internal static A.Outline Build(PresentationShape source)
    {
        var requested = FromWire(source);
        var outline = new A.Outline { Width = checked((int)requested.WidthEmu) };
        AppendProfile(outline, requested);
        return outline;
    }

    internal static void Apply(P.ShapeProperties properties, PresentationShape source)
    {
        var requested = FromWire(source);
        var outline = properties.GetFirstChild<A.Outline>();
        if (!TryRead(outline, out var current))
            throw new CodecException("unsupported_presentation_line", "The imported presentation shape line is outside the editable RGB/preset-dash profile.");
        if (current == requested) return;

        if (outline is null)
        {
            outline = new A.Outline();
            OpenXmlElement? anchor = properties.ChildElements.LastOrDefault(child => child is A.NoFill or A.SolidFill);
            anchor ??= properties.GetFirstChild<A.CustomGeometry>();
            anchor ??= properties.GetFirstChild<A.PresetGeometry>();
            anchor ??= properties.GetFirstChild<A.Transform2D>();
            if (anchor is null) properties.PrependChild(outline);
            else properties.InsertAfter(outline, anchor);
        }
        outline.Width = checked((int)requested.WidthEmu);
        outline.RemoveAllChildren();
        AppendProfile(outline, requested);
    }

    internal static void ScrubModeledContent(A.Outline outline)
    {
        outline.Width = 0;
        foreach (var child in outline.ChildElements.Where(child => child is A.NoFill or A.SolidFill or A.PresetDash).ToArray())
            child.Remove();
    }

    private static Profile FromWire(PresentationShape source)
    {
        var style = NormalizeStyle(source.LineStyle, source.LineRgb);
        var rgb = style == "none" ? string.Empty : PptxColor.Normalize(source.LineRgb);
        return new Profile(rgb, source.LineWidthEmu, style);
    }

    private static string NormalizeStyle(string style, string rgb) =>
        string.IsNullOrWhiteSpace(style)
            ? string.IsNullOrWhiteSpace(rgb) ? "none" : "solid"
            : style;

    private static void AppendProfile(A.Outline outline, Profile profile)
    {
        if (profile.Style == "none")
        {
            outline.Append(new A.NoFill());
            return;
        }
        outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = profile.Rgb }));
        outline.Append(new A.PresetDash { Val = profile.Style switch
        {
            "dashed" => A.PresetLineDashValues.Dash,
            "dotted" => A.PresetLineDashValues.Dot,
            "dash-dot" => A.PresetLineDashValues.DashDot,
            "dash-dot-dot" => A.PresetLineDashValues.LargeDashDotDot,
            _ => A.PresetLineDashValues.Solid,
        } });
    }

    private static string FallbackRgb(A.SolidFill? solid)
    {
        var color = solid?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
        return color is { Length: 6 } && color.All(Uri.IsHexDigit) ? color.ToUpperInvariant() : string.Empty;
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }
}
