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
        string Scheme,
        uint? OpacityThousandthPercent,
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

    // Keep the source-bound native leaf vocabulary in semantic terms while
    // retaining the exact DrawingML token for the final token splice.
    internal static bool TryReadPresetDash(A.PresetDash? source, out string style)
    {
        style = string.Empty;
        if (source is null || source.ChildElements.Any() || !HasOnlyAttributes(source, "val") || source.Val?.Value is not { } value)
            return false;
        style = value.Equals(A.PresetLineDashValues.Solid) ? "solid" :
            value.Equals(A.PresetLineDashValues.Dash) ? "dashed" :
            value.Equals(A.PresetLineDashValues.Dot) ? "dotted" :
            value.Equals(A.PresetLineDashValues.DashDot) ? "dash-dot" :
            value.Equals(A.PresetLineDashValues.LargeDashDotDot) ? "dash-dot-dot" :
            string.Empty;
        return style.Length > 0;
    }

    internal static bool TryReadPresetDashValue(string? value, out string style)
    {
        var normalized = value?.Trim().ToLowerInvariant();
        style = normalized switch
        {
            "solid" => "solid",
            "dash" => "dashed",
            "dot" => "dotted",
            "dashdot" => "dash-dot",
            "lgdashdotdot" => "dash-dot-dot",
            _ => string.Empty,
        };
        return style.Length > 0;
    }

    internal static bool TryPresetDashToken(string style, out string token)
    {
        token = style switch
        {
            "solid" => "solid",
            "dashed" => "dash",
            "dotted" => "dot",
            "dash-dot" => "dashDot",
            "dash-dot-dot" => "lgDashDotDot",
            _ => string.Empty,
        };
        return token.Length > 0;
    }

    internal static bool TryReadCapValue(string? value, out string cap)
    {
        cap = value?.Trim().ToLowerInvariant() switch
        {
            "flat" => "flat",
            "rnd" or "round" => "round",
            "sq" or "square" => "square",
            _ => string.Empty,
        };
        return cap.Length > 0;
    }

    internal static bool TryReadCap(A.Outline? outline, out string cap)
    {
        cap = string.Empty;
        if (outline is null) return false;
        var attributes = outline.GetAttributes()
            .Where(attribute => attribute.LocalName == "cap")
            .ToArray();
        return attributes.Length == 1 && HasOnlyAttributes(outline, "w", "cap", "cmpd", "algn") &&
            TryReadCapValue(attributes[0].Value, out cap);
    }

    internal static bool TryCapToken(string cap, out string token)
    {
        token = cap switch
        {
            "flat" => "flat",
            "round" => "rnd",
            "square" => "sq",
            _ => string.Empty,
        };
        return token.Length > 0;
    }

    // Source-bound joins are intentionally narrower than the full typed line
    // profile: only one bare DrawingML join element is safe to token-splice.
    // A miter limit, extra attributes, or child markup stays opaque.
    internal static bool TryReadJoinLeaf(A.Outline? outline, out string join)
    {
        join = string.Empty;
        if (outline is null) return false;
        var joins = outline.ChildElements.Where(child => child is A.Round or A.LineJoinBevel or A.Miter).ToArray();
        if (joins.Length != 1) return false;
        var candidate = joins[0];
        if (candidate.ChildElements.Any() || !HasOnlyAttributes(candidate)) return false;
        join = candidate switch
        {
            A.Round => "round",
            A.LineJoinBevel => "bevel",
            A.Miter => "miter",
            _ => string.Empty,
        };
        return join.Length > 0;
    }

    internal static bool TryJoinToken(string join, out string token)
    {
        token = join switch
        {
            "round" => "round",
            "bevel" => "bevel",
            "miter" => "miter",
            _ => string.Empty,
        };
        return token.Length > 0;
    }

    // Source-bound arrow leaves only splice an existing explicit endpoint
    // type. Width/length attributes are retained and must already be
    // canonical, so an irregular endpoint stays opaque.
    internal static bool TryReadArrowType(OpenXmlElement? source, out string type)
    {
        type = string.Empty;
        if (source is not (A.HeadEnd or A.TailEnd) || source.ChildElements.Any() ||
            !HasOnlyAttributes(source, "type", "w", "len") || source.GetAttributes().Count(attribute => attribute.LocalName == "type") != 1) return false;
        var sourceType = source is A.HeadEnd head ? head.Type?.Value : ((A.TailEnd)source).Type?.Value;
        var sourceWidth = source is A.HeadEnd headWidth ? headWidth.Width?.Value : ((A.TailEnd)source).Width?.Value;
        var sourceLength = source is A.HeadEnd headLength ? headLength.Length?.Value : ((A.TailEnd)source).Length?.Value;
        if (!TryArrow(sourceType, out var parsed) || !TryEndWidth(sourceWidth, out _) || !TryEndLength(sourceLength, out _)) return false;
        type = parsed.Length == 0 ? "none" : parsed;
        return true;
    }

    internal static bool TryArrowTypeToken(string type, out string token)
    {
        token = type switch
        {
            "none" => "none",
            "triangle" => "triangle",
            "stealth" => "stealth",
            "diamond" => "diamond",
            "oval" => "oval",
            "arrow" => "arrow",
            _ => string.Empty,
        };
        return token.Length > 0;
    }

    internal static bool TryArrowSizeToken(string size) => size is "sm" or "med" or "lg";

    internal static bool TryRead(A.Outline? outline, out Profile profile)
    {
        if (outline is null)
        {
            profile = EmptyProfile();
            return true;
        }

        profile = null!;
        var width = outline.Width?.Value ?? 0;
        if (width < 0 || width > int.MaxValue ||
            !HasOnlyAttributes(outline, "w", "cap", "cmpd", "algn") ||
            !TryCap(outline.CapType?.Value, out var cap))
            return false;

        // These are the canonical single-line/centered defaults. A number of
        // exporters write them explicitly on otherwise ordinary connectors;
        // accepting only these values keeps the editable profile bounded while
        // preserving the common native representation.
        if (outline.CompoundLineType?.Value is { } compound &&
            !compound.Equals(A.CompoundLineValues.Single)) return false;
        if (outline.Alignment?.Value is { } alignment &&
            !alignment.Equals(A.PenAlignmentValues.Center)) return false;

        var noFills = outline.Elements<A.NoFill>().ToArray();
        var solidFills = outline.Elements<A.SolidFill>().ToArray();
        // The rendered result of an a:ln without EG_LineFillProperties is not
        // stable enough to expose as editable. Require one explicit noFill or
        // one direct RGB/theme solidFill.
        if (noFills.Length + solidFills.Length != 1) return false;

        var rgb = string.Empty;
        var scheme = string.Empty;
        uint? opacity = null;
        var noFill = noFills.Length == 1;
        if (noFill && (noFills[0].ChildElements.Any() || !HasOnlyAttributes(noFills[0]))) return false;
        if (solidFills.Length == 1)
        {
            var solid = solidFills[0];
            if (!HasOnlyAttributes(solid) ||
                !PptxColor.TryDirectSolidRgbWithOpacity(solid, out rgb, out opacity) &&
                !PptxColor.TryDirectSolidSchemeWithOpacity(solid, out scheme, out opacity)) return false;
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

        profile = new Profile(rgb, scheme, opacity, width, style, cap, join,
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
        target.LineScheme = FallbackScheme(outline?.GetFirstChild<A.SolidFill>());
        var width = outline?.Width?.Value ?? 0;
        target.LineWidthEmu = width is >= 0 and <= int.MaxValue ? width : 0;
        target.LineStyle = target.LineRgb.Length > 0 || target.LineScheme.Length > 0 ? "solid" : "none";
        target.StartArrow = target.EndArrow = target.StartArrowWidth = target.StartArrowLength =
            target.EndArrowWidth = target.EndArrowLength = target.LineCap = target.LineJoin = string.Empty;
    }

    internal static void CopyTo(Profile source, PresentationConnector target)
    {
        target.LineRgb = source.Rgb;
        target.LineScheme = source.Scheme;
        if (source.OpacityThousandthPercent is { } opacity) target.LineOpacityThousandthPercent = opacity;
        else target.ClearLineOpacityThousandthPercent();
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

    internal static void NormalizeSemantics(PresentationShape source)
    {
        source.LineStyle = NormalizeStyle(source.LineStyle, source.LineRgb, source.LineScheme);
        source.LineRgb = string.IsNullOrWhiteSpace(source.LineRgb) ? string.Empty : PptxColor.Normalize(source.LineRgb);
        source.LineScheme = string.IsNullOrWhiteSpace(source.LineScheme) ? string.Empty : PptxColor.NormalizeScheme(source.LineScheme);
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
            if (!string.IsNullOrWhiteSpace(source.Rgb) || !string.IsNullOrWhiteSpace(source.Scheme) || source.OpacityThousandthPercent is not null)
                throw new CodecException(invalidCode, $"{subject} cannot combine line style none with a color.");
            return source with { Rgb = string.Empty, Scheme = string.Empty };
        }
        if (string.IsNullOrWhiteSpace(source.Rgb) == string.IsNullOrWhiteSpace(source.Scheme))
            throw new CodecException(invalidCode, $"{subject} line style {source.Style} requires exactly one RGB or theme color.");
        if (source.OpacityThousandthPercent is > 100_000)
            throw new CodecException(invalidCode, $"{subject} has an invalid line opacity.");
        return source with
        {
            Rgb = string.IsNullOrWhiteSpace(source.Rgb) ? string.Empty : PptxColor.Normalize(source.Rgb),
            Scheme = string.IsNullOrWhiteSpace(source.Scheme) ? string.Empty : PptxColor.NormalizeScheme(source.Scheme),
        };
    }

    private static Profile FromWire(PresentationShape source) => new(
        source.LineRgb,
        source.LineScheme,
        source.HasLineOpacityThousandthPercent ? source.LineOpacityThousandthPercent : null,
        source.LineWidthEmu,
        NormalizeStyle(source.LineStyle, source.LineRgb, source.LineScheme),
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
        source.LineScheme,
        source.HasLineOpacityThousandthPercent ? source.LineOpacityThousandthPercent : null,
        source.LineWidthEmu,
        NormalizeStyle(source.LineStyle, source.LineRgb, source.LineScheme),
        source.LineCap,
        source.LineJoin,
        source.StartArrow,
        source.StartArrowWidth,
        source.StartArrowLength,
        source.EndArrow,
        source.EndArrowWidth,
        source.EndArrowLength);

    private static string NormalizeStyle(string style, string rgb, string scheme = "") =>
        string.IsNullOrWhiteSpace(style)
            ? string.IsNullOrWhiteSpace(rgb) && string.IsNullOrWhiteSpace(scheme) ? "none" : "solid"
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
            OpenXmlElement color = source.Scheme.Length > 0
                ? new A.SchemeColor { Val = PptxColor.SchemeValue(source.Scheme) }
                : new A.RgbColorModelHex { Val = PptxColor.Normalize(source.Rgb) };
            if (source.OpacityThousandthPercent is { } opacity)
                color.Append(new A.Alpha { Val = checked((int)opacity) });
            outline.Append(new A.SolidFill(color));
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
        // `type="none"` has no rendered head/tail. Some native exporters
        // still serialize the default width/length attributes alongside it;
        // those values are inert and cannot affect the resulting line.
        if (type.Length == 0)
        {
            width = length = string.Empty;
            return true;
        }
        return true;
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
        target.LineScheme = source.Scheme;
        if (source.OpacityThousandthPercent is { } opacity) target.LineOpacityThousandthPercent = opacity;
        else target.ClearLineOpacityThousandthPercent();
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

    private static Profile EmptyProfile() => new(string.Empty, string.Empty, null, 0, "none", string.Empty, string.Empty,
        string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

    private static bool IsProfileChild(OpenXmlElement child) =>
        child is A.NoFill or A.SolidFill or A.PresetDash or A.Round or A.LineJoinBevel or A.Miter or A.HeadEnd or A.TailEnd;

    private static string FallbackRgb(A.SolidFill? solid)
    {
        var color = solid?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
        return color is { Length: 6 } && color.All(Uri.IsHexDigit) ? color.ToUpperInvariant() : string.Empty;
    }

    private static string FallbackScheme(A.SolidFill? solid) => PptxColor.SolidScheme(solid);

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }
}
