using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the bounded direct DrawingML a:ln profile shared by ordinary p:sp
// outlines and p:cxnSp connectors. Geometry, connector targets, routing, and
// transforms stay with their element codecs; this module owns only paint,
// width, preset dash, cap, join, and line ends.
internal static class PptxLineStyleCodec
{
    internal sealed record Profile(
        string Rgb,
        long WidthEmu,
        string Style,
        string Cap,
        string Join,
        string StartArrow,
        string StartArrowWidth,
        string StartArrowLength,
        string EndArrow,
        string EndArrowWidth,
        string EndArrowLength);

    private static readonly IReadOnlySet<string> Styles = new HashSet<string>(StringComparer.Ordinal)
    {
        "solid", "dashed", "dotted", "dash-dot", "dash-dot-dot", "none",
    };
    private static readonly IReadOnlySet<string> ArrowTypes = new HashSet<string>(StringComparer.Ordinal)
    {
        "", "triangle", "stealth", "diamond", "oval", "arrow",
    };
    private static readonly IReadOnlySet<string> EndSizes = new HashSet<string>(StringComparer.Ordinal)
    {
        "", "sm", "med", "lg",
    };
    private static readonly IReadOnlySet<string> Caps = new HashSet<string>(StringComparer.Ordinal)
    {
        "", "flat", "round", "square",
    };
    private static readonly IReadOnlySet<string> Joins = new HashSet<string>(StringComparer.Ordinal)
    {
        "", "round", "bevel", "miter",
    };

    internal static bool TryRead(A.Outline? outline, out Profile profile)
    {
        if (outline is null)
        {
            profile = EmptyProfile();
            return true;
        }

        profile = null!;
        var width = outline.Width?.Value ?? 0;
        if (width < 0 || width > int.MaxValue || outline.CompoundLineType is not null ||
            outline.Alignment is not null || !HasOnlyAttributes(outline, "w", "cap") ||
            !TryCap(outline.CapType?.Value, out var cap))
            return false;

        var noFills = outline.Elements<A.NoFill>().ToArray();
        var solidFills = outline.Elements<A.SolidFill>().ToArray();
        // The rendered result of an a:ln without EG_LineFillProperties is not
        // stable enough to expose as editable. Require one explicit noFill or
        // one direct sRGB solidFill.
        if (noFills.Length + solidFills.Length != 1) return false;

        var rgb = string.Empty;
        var noFill = noFills.Length == 1;
        if (noFill && (noFills[0].ChildElements.Any() || !HasOnlyAttributes(noFills[0]))) return false;
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
        var dash = dashes.SingleOrDefault()?.Val?.Value;
        string style;
        if (noFill)
        {
            if (dash is not null) return false;
            style = "none";
        }
        else
        {
            style = dash is null || dash.Value.Equals(A.PresetLineDashValues.Solid) ? "solid" :
                dash.Value.Equals(A.PresetLineDashValues.Dash) ? "dashed" :
                dash.Value.Equals(A.PresetLineDashValues.Dot) ? "dotted" :
                dash.Value.Equals(A.PresetLineDashValues.DashDot) ? "dash-dot" :
                dash.Value.Equals(A.PresetLineDashValues.LargeDashDotDot) ? "dash-dot-dot" : string.Empty;
            if (style.Length == 0) return false;
        }

        if (!TryJoin(outline, out var join)) return false;
        var heads = outline.Elements<A.HeadEnd>().ToArray();
        var tails = outline.Elements<A.TailEnd>().ToArray();
        if (heads.Length > 1 || tails.Length > 1 ||
            !TryLineEnd(heads.SingleOrDefault(), out var startArrow, out var startArrowWidth, out var startArrowLength) ||
            !TryLineEnd(tails.SingleOrDefault(), out var endArrow, out var endArrowWidth, out var endArrowLength) ||
            outline.ChildElements.Any(child => child is not A.NoFill and not A.SolidFill and not A.PresetDash and
                not A.Round and not A.LineJoinBevel and not A.Miter and not A.HeadEnd and not A.TailEnd))
            return false;

        profile = new Profile(rgb, width, style, cap, join,
            startArrow, startArrowWidth, startArrowLength,
            endArrow, endArrowWidth, endArrowLength);
        return true;
    }

    internal static void ReadForProjection(A.Outline? outline, PresentationShape target)
    {
        if (TryRead(outline, out var profile))
        {
            CopyTo(profile, target);
            return;
        }

        // The enclosing element remains source-bound. Keep the historic RGB
        // inspection projection without granting editability.
        target.LineRgb = FallbackRgb(outline?.GetFirstChild<A.SolidFill>());
        var width = outline?.Width?.Value ?? 0;
        target.LineWidthEmu = width is >= 0 and <= int.MaxValue ? width : 0;
        target.LineStyle = target.LineRgb.Length > 0 ? "solid" : "none";
        target.StartArrow = target.EndArrow = target.StartArrowWidth = target.StartArrowLength =
            target.EndArrowWidth = target.EndArrowLength = target.LineCap = target.LineJoin = string.Empty;
    }

    internal static void CopyTo(Profile source, PresentationConnector target)
    {
        target.LineRgb = source.Rgb;
        target.LineWidthEmu = source.WidthEmu;
        target.LineStyle = source.Style;
        target.LineCap = source.Cap;
        target.LineJoin = source.Join;
        target.StartArrow = source.StartArrow;
        target.StartArrowWidth = source.StartArrowWidth;
        target.StartArrowLength = source.StartArrowLength;
        target.EndArrow = source.EndArrow;
        target.EndArrowWidth = source.EndArrowWidth;
        target.EndArrowLength = source.EndArrowLength;
    }

    internal static void Validate(PresentationShape source, string shapeId)
    {
        _ = ValidateProfile(FromWire(source), $"Presentation shape {shapeId}",
            "invalid_presentation_line", "unsupported_presentation_line");
        if (!string.Equals(source.Geometry, "line", StringComparison.Ordinal) &&
            (source.StartArrow.Length > 0 || source.EndArrow.Length > 0))
            throw new CodecException("unsupported_presentation_line", $"Presentation shape {shapeId} arrowheads require geometry line.");
    }

    internal static void Validate(PresentationConnector source, string elementId) =>
        _ = ValidateProfile(FromWire(source), $"Presentation connector {elementId}",
            "invalid_presentation_connector", "unsupported_presentation_connector");

    internal static A.Outline Build(PresentationShape source) => Build(FromWire(source), emitSolidDash: true);

    internal static A.Outline Build(PresentationConnector source) => Build(FromWire(source), emitSolidDash: false);

    internal static void Apply(P.ShapeProperties properties, PresentationShape source)
    {
        var requested = FromWire(source);
        var outline = properties.GetFirstChild<A.Outline>();
        if (!TryRead(outline, out var current))
            throw new CodecException("unsupported_presentation_line", "The imported presentation shape line is outside the editable direct a:ln profile.");
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
        ReplaceProfile(outline, requested, emitSolidDash: true);
    }

    internal static void ScrubModeledContent(A.Outline outline)
    {
        outline.Width = 0;
        outline.CapType = null;
        foreach (var child in outline.ChildElements.Where(IsProfileChild).ToArray()) child.Remove();
    }

    private static Profile ValidateProfile(Profile source, string subject, string invalidCode, string unsupportedCode)
    {
        if (source.WidthEmu < 0 || source.WidthEmu > int.MaxValue)
            throw new CodecException(invalidCode, $"{subject} has an invalid line width.");
        if (!Styles.Contains(source.Style) || !ArrowTypes.Contains(source.StartArrow) || !ArrowTypes.Contains(source.EndArrow) ||
            !EndSizes.Contains(source.StartArrowWidth) || !EndSizes.Contains(source.StartArrowLength) ||
            !EndSizes.Contains(source.EndArrowWidth) || !EndSizes.Contains(source.EndArrowLength) ||
            !Caps.Contains(source.Cap) || !Joins.Contains(source.Join))
            throw new CodecException(unsupportedCode, $"{subject} uses unsupported line styling.");
        if (source.StartArrow.Length == 0 && (source.StartArrowWidth.Length > 0 || source.StartArrowLength.Length > 0) ||
            source.EndArrow.Length == 0 && (source.EndArrowWidth.Length > 0 || source.EndArrowLength.Length > 0))
            throw new CodecException(invalidCode, $"{subject} has incomplete line-end state.");
        if (source.Style == "none")
        {
            if (!string.IsNullOrWhiteSpace(source.Rgb))
                throw new CodecException(invalidCode, $"{subject} cannot combine line style none with a color.");
            return source with { Rgb = string.Empty };
        }
        if (string.IsNullOrWhiteSpace(source.Rgb))
            throw new CodecException(invalidCode, $"{subject} line style {source.Style} requires an RGB color.");
        return source with { Rgb = PptxColor.Normalize(source.Rgb) };
    }

    private static Profile FromWire(PresentationShape source) => new(
        source.LineRgb,
        source.LineWidthEmu,
        NormalizeStyle(source.LineStyle, source.LineRgb),
        source.LineCap,
        source.LineJoin,
        source.StartArrow,
        source.StartArrowWidth,
        source.StartArrowLength,
        source.EndArrow,
        source.EndArrowWidth,
        source.EndArrowLength);

    private static Profile FromWire(PresentationConnector source) => new(
        source.LineRgb,
        source.LineWidthEmu,
        NormalizeStyle(source.LineStyle, source.LineRgb),
        source.LineCap,
        source.LineJoin,
        source.StartArrow,
        source.StartArrowWidth,
        source.StartArrowLength,
        source.EndArrow,
        source.EndArrowWidth,
        source.EndArrowLength);

    private static string NormalizeStyle(string style, string rgb) =>
        string.IsNullOrWhiteSpace(style)
            ? string.IsNullOrWhiteSpace(rgb) ? "none" : "solid"
            : style;

    private static A.Outline Build(Profile source, bool emitSolidDash)
    {
        var outline = new A.Outline();
        ReplaceProfile(outline, source, emitSolidDash);
        return outline;
    }

    private static void ReplaceProfile(A.Outline outline, Profile source, bool emitSolidDash)
    {
        outline.Width = checked((int)source.WidthEmu);
        outline.CapType = source.Cap.Length == 0 ? null : source.Cap switch
        {
            "round" => A.LineCapValues.Round,
            "square" => A.LineCapValues.Square,
            _ => A.LineCapValues.Flat,
        };
        outline.RemoveAllChildren();
        if (source.Style == "none") outline.Append(new A.NoFill());
        else
        {
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(source.Rgb) }));
            if (emitSolidDash || source.Style != "solid") outline.Append(new A.PresetDash { Val = DashValue(source.Style) });
        }
        if (source.Join.Length > 0) outline.Append(source.Join switch
        {
            "round" => new A.Round(),
            "bevel" => new A.LineJoinBevel(),
            _ => new A.Miter(),
        });
        if (source.StartArrow.Length > 0) outline.Append(HeadEnd(source.StartArrow, source.StartArrowWidth, source.StartArrowLength));
        if (source.EndArrow.Length > 0) outline.Append(TailEnd(source.EndArrow, source.EndArrowWidth, source.EndArrowLength));
    }

    private static A.PresetLineDashValues DashValue(string style) => style switch
    {
        "dashed" => A.PresetLineDashValues.Dash,
        "dotted" => A.PresetLineDashValues.Dot,
        "dash-dot" => A.PresetLineDashValues.DashDot,
        "dash-dot-dot" => A.PresetLineDashValues.LargeDashDotDot,
        _ => A.PresetLineDashValues.Solid,
    };

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
            if (round.ChildElements.Any() || !HasOnlyAttributes(round)) return false;
            join = "round";
        }
        else if (joins.SingleOrDefault() is A.LineJoinBevel bevel)
        {
            if (bevel.ChildElements.Any() || !HasOnlyAttributes(bevel)) return false;
            join = "bevel";
        }
        else if (joins.SingleOrDefault() is A.Miter miter)
        {
            if (miter.ChildElements.Any() || !HasOnlyAttributes(miter) || miter.Limit is not null) return false;
            join = "miter";
        }
        return true;
    }

    private static bool TryLineEnd(A.HeadEnd? source, out string type, out string width, out string length) =>
        TryLineEnd(source, source?.Type?.Value, source?.Width?.Value, source?.Length?.Value, out type, out width, out length);

    private static bool TryLineEnd(A.TailEnd? source, out string type, out string width, out string length) =>
        TryLineEnd(source, source?.Type?.Value, source?.Width?.Value, source?.Length?.Value, out type, out width, out length);

    private static bool TryLineEnd(
        OpenXmlElement? source,
        A.LineEndValues? sourceType,
        A.LineEndWidthValues? sourceWidth,
        A.LineEndLengthValues? sourceLength,
        out string type,
        out string width,
        out string length)
    {
        type = width = length = string.Empty;
        if (source is null) return true;
        if (source.ChildElements.Any() || !HasOnlyAttributes(source, "type", "w", "len") ||
            !TryArrow(sourceType, out type) || !TryEndWidth(sourceWidth, out width) || !TryEndLength(sourceLength, out length))
            return false;
        return type.Length > 0 || width.Length == 0 && length.Length == 0;
    }

    private static bool TryArrow(A.LineEndValues? source, out string arrow)
    {
        arrow = string.Empty;
        if (source is null || source.Value.Equals(A.LineEndValues.None)) return true;
        if (source.Value.Equals(A.LineEndValues.Triangle)) arrow = "triangle";
        else if (source.Value.Equals(A.LineEndValues.Stealth)) arrow = "stealth";
        else if (source.Value.Equals(A.LineEndValues.Diamond)) arrow = "diamond";
        else if (source.Value.Equals(A.LineEndValues.Oval)) arrow = "oval";
        else if (source.Value.Equals(A.LineEndValues.Arrow)) arrow = "arrow";
        else return false;
        return true;
    }

    private static bool TryEndWidth(A.LineEndWidthValues? source, out string width)
    {
        width = string.Empty;
        if (source is null) return true;
        if (source.Value.Equals(A.LineEndWidthValues.Small)) width = "sm";
        else if (source.Value.Equals(A.LineEndWidthValues.Medium)) width = "med";
        else if (source.Value.Equals(A.LineEndWidthValues.Large)) width = "lg";
        else return false;
        return true;
    }

    private static bool TryEndLength(A.LineEndLengthValues? source, out string length)
    {
        length = string.Empty;
        if (source is null) return true;
        if (source.Value.Equals(A.LineEndLengthValues.Small)) length = "sm";
        else if (source.Value.Equals(A.LineEndLengthValues.Medium)) length = "med";
        else if (source.Value.Equals(A.LineEndLengthValues.Large)) length = "lg";
        else return false;
        return true;
    }

    private static A.HeadEnd HeadEnd(string type, string width, string length)
    {
        var output = new A.HeadEnd { Type = LineEndType(type) };
        if (width.Length > 0) output.Width = LineEndWidth(width);
        if (length.Length > 0) output.Length = LineEndLength(length);
        return output;
    }

    private static A.TailEnd TailEnd(string type, string width, string length)
    {
        var output = new A.TailEnd { Type = LineEndType(type) };
        if (width.Length > 0) output.Width = LineEndWidth(width);
        if (length.Length > 0) output.Length = LineEndLength(length);
        return output;
    }

    private static A.LineEndValues LineEndType(string type) => type switch
    {
        "stealth" => A.LineEndValues.Stealth,
        "diamond" => A.LineEndValues.Diamond,
        "oval" => A.LineEndValues.Oval,
        "arrow" => A.LineEndValues.Arrow,
        _ => A.LineEndValues.Triangle,
    };

    private static A.LineEndWidthValues LineEndWidth(string width) => width switch
    {
        "sm" => A.LineEndWidthValues.Small,
        "lg" => A.LineEndWidthValues.Large,
        _ => A.LineEndWidthValues.Medium,
    };

    private static A.LineEndLengthValues LineEndLength(string length) => length switch
    {
        "sm" => A.LineEndLengthValues.Small,
        "lg" => A.LineEndLengthValues.Large,
        _ => A.LineEndLengthValues.Medium,
    };

    private static void CopyTo(Profile source, PresentationShape target)
    {
        target.LineRgb = source.Rgb;
        target.LineWidthEmu = source.WidthEmu;
        target.LineStyle = source.Style;
        target.LineCap = source.Cap;
        target.LineJoin = source.Join;
        target.StartArrow = source.StartArrow;
        target.StartArrowWidth = source.StartArrowWidth;
        target.StartArrowLength = source.StartArrowLength;
        target.EndArrow = source.EndArrow;
        target.EndArrowWidth = source.EndArrowWidth;
        target.EndArrowLength = source.EndArrowLength;
    }

    private static Profile EmptyProfile() => new(string.Empty, 0, "none", string.Empty, string.Empty,
        string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

    private static bool IsProfileChild(OpenXmlElement child) =>
        child is A.NoFill or A.SolidFill or A.PresetDash or A.Round or A.LineJoinBevel or A.Miter or A.HeadEnd or A.TailEnd;

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
