using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

internal static class PptxColor
{
    private static readonly IReadOnlyDictionary<string, A.SchemeColorValues> SchemeColors =
        new Dictionary<string, A.SchemeColorValues>(StringComparer.Ordinal)
        {
            ["bg1"] = A.SchemeColorValues.Background1,
            ["tx1"] = A.SchemeColorValues.Text1,
            ["bg2"] = A.SchemeColorValues.Background2,
            ["tx2"] = A.SchemeColorValues.Text2,
            ["accent1"] = A.SchemeColorValues.Accent1,
            ["accent2"] = A.SchemeColorValues.Accent2,
            ["accent3"] = A.SchemeColorValues.Accent3,
            ["accent4"] = A.SchemeColorValues.Accent4,
            ["accent5"] = A.SchemeColorValues.Accent5,
            ["accent6"] = A.SchemeColorValues.Accent6,
            ["hlink"] = A.SchemeColorValues.Hyperlink,
            ["folHlink"] = A.SchemeColorValues.FollowedHyperlink,
            ["dk1"] = A.SchemeColorValues.Dark1,
            ["lt1"] = A.SchemeColorValues.Light1,
            ["dk2"] = A.SchemeColorValues.Dark2,
            ["lt2"] = A.SchemeColorValues.Light2,
        };

    internal static string SolidRgb(A.SolidFill? fill) =>
        fill?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value ?? string.Empty;

    // Source-bound run color edits are limited to a bare, direct RGB paint.
    // Effects, alpha transforms, scheme colors, and extra children remain
    // source-owned so a native-leaf operation cannot accidentally normalize
    // the surrounding run properties.
    internal static bool TryDirectSolidRgb(A.SolidFill? fill, out string rgb)
    {
        if (!TryDirectSolidRgbWithOpacity(fill, out rgb, out var opacity) || opacity is not null)
        {
            rgb = string.Empty;
            return false;
        }
        return true;
    }

    // Semantic projection may own one direct alpha child while native-leaf
    // color edits remain intentionally limited to the bare form above.
    internal static bool TryDirectSolidRgbWithOpacity(A.SolidFill? fill, out string rgb, out uint? opacity)
    {
        rgb = string.Empty;
        opacity = null;
        if (fill is null || fill.ChildElements.Count != 1 || !HasOnlyAttributes(fill)) return false;
        if (fill.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") || color.Val?.Value is not { Length: 6 } value ||
            !value.All(Uri.IsHexDigit) || !TryDirectOpacity(color, out opacity)) return false;
        rgb = value.ToUpperInvariant();
        return true;
    }

    // Source-bound run color edits can also preserve a bare theme token. Keep
    // this strict: scheme colors with transforms or extra attributes remain
    // source-owned rather than being rebuilt from a lossy semantic value.
    internal static bool TryDirectSolidScheme(A.SolidFill? fill, out string scheme)
    {
        if (!TryDirectSolidSchemeWithOpacity(fill, out scheme, out var opacity) || opacity is not null ||
            fill?.FirstChild is not A.SchemeColor color || color.ChildElements.Count != 0)
        {
            scheme = string.Empty;
            return false;
        }
        return true;
    }

    internal static bool TryDirectSolidSchemeWithOpacity(A.SolidFill? fill, out string scheme, out uint? opacity)
    {
        scheme = string.Empty;
        opacity = null;
        if (fill is null || fill.ChildElements.Count != 1 || !HasOnlyAttributes(fill)) return false;
        if (fill.FirstChild is not A.SchemeColor color ||
            !HasOnlyAttributes(color, "val") || color.Val?.Value is not { } value ||
            !TrySchemeToken(value, out var token) || !TryDirectSchemeTransforms(color, out opacity)) return false;
        scheme = token;
        return true;
    }

    internal static A.SolidFill BuildSolidRgb(string value, uint? opacity)
    {
        var color = new A.RgbColorModelHex { Val = Normalize(value) };
        AppendOpacity(color, opacity);
        return new A.SolidFill(color);
    }

    internal static A.SolidFill BuildSolidScheme(string value, uint? opacity)
    {
        var color = new A.SchemeColor { Val = SchemeValue(value) };
        AppendOpacity(color, opacity);
        return new A.SolidFill(color);
    }

    internal static string SolidScheme(A.SolidFill? fill)
    {
        if (fill is null || fill.ChildElements.Count != 1 ||
            fill.FirstChild is not A.SchemeColor scheme ||
            scheme.ChildElements.Count != 0 || !HasOnlyAttributes(scheme, "val") ||
            scheme.Val?.Value is not { } value || !TrySchemeToken(value, out var token))
            return string.Empty;
        return token;
    }

    // Semantic projection may carry a single bounded alpha child alongside a
    // direct theme token. Keep the strict bare helper above for callers that
    // intentionally require an unmodified scheme color, while allowing the
    // shape projection to preserve this safe source-bound form.
    internal static string SolidSchemeWithOpacity(A.SolidFill? fill)
    {
        return TryDirectSolidSchemeWithOpacity(fill, out var scheme, out _)
            ? scheme
            : string.Empty;
    }

    private static bool HasOnlyAttributes(DocumentFormat.OpenXml.OpenXmlElement element, params string[] allowed)
    {
        var accepted = allowed.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => accepted.Contains(attribute.LocalName));
    }

    private static bool TryDirectOpacity(DocumentFormat.OpenXml.OpenXmlElement color, out uint? opacity)
    {
        opacity = null;
        if (color.ChildElements.Count == 0) return true;
        if (color.ChildElements.Count != 1 || color.FirstChild is not A.Alpha alpha ||
            alpha.ChildElements.Count != 0 || !HasOnlyAttributes(alpha, "val") ||
            alpha.Val?.Value is not { } value || value is < 0 or > 100_000) return false;
        opacity = checked((uint)value);
        return true;
    }

    // Theme colors in imported decks sometimes carry luminance transforms in
    // addition to alpha (for example lumMod/lumOff/alpha). These transforms
    // are retained as source-owned XML; accepting this small, fully bounded
    // profile lets a native leaf change only alpha or the theme token without
    // reconstructing the visual paint.
    private static bool TryDirectSchemeTransforms(A.SchemeColor color, out uint? opacity)
    {
        opacity = null;
        var alpha = color.Elements<A.Alpha>().ToArray();
        var luminanceModulation = color.Elements<A.LuminanceModulation>().ToArray();
        var luminanceOffset = color.Elements<A.LuminanceOffset>().ToArray();
        if (alpha.Length > 1 || luminanceModulation.Length > 1 || luminanceOffset.Length > 1 ||
            color.ChildElements.Any(child => child is not A.Alpha and not A.LuminanceModulation and not A.LuminanceOffset))
            return false;
        if (alpha.Length == 1)
        {
            var value = alpha[0].Val?.Value;
            if (alpha[0].ChildElements.Count != 0 || !HasOnlyAttributes(alpha[0], "val") || value is not (>= 0 and <= 100_000))
                return false;
            opacity = checked((uint)value);
        }
        if (luminanceModulation.Length == 1)
        {
            var value = luminanceModulation[0].Val?.Value;
            if (luminanceModulation[0].ChildElements.Count != 0 || !HasOnlyAttributes(luminanceModulation[0], "val") || value is not (>= 0 and <= 100_000))
                return false;
        }
        if (luminanceOffset.Length == 1)
        {
            var value = luminanceOffset[0].Val?.Value;
            if (luminanceOffset[0].ChildElements.Count != 0 || !HasOnlyAttributes(luminanceOffset[0], "val") || value is not (>= -100_000 and <= 100_000))
                return false;
        }
        return true;
    }

    private static void AppendOpacity(DocumentFormat.OpenXml.OpenXmlElement color, uint? opacity)
    {
        if (opacity is null) return;
        if (opacity > 100_000)
            throw new CodecException("invalid_presentation_color", "Presentation color opacity must be at most 100000 thousandths of a percent.");
        color.Append(new A.Alpha { Val = checked((int)opacity.Value) });
    }

    internal static string Normalize(string value)
    {
        var rgb = value.Trim().TrimStart('#').ToUpperInvariant();
        if (rgb.Length != 6 || rgb.Any(character => !Uri.IsHexDigit(character)))
            throw new CodecException("invalid_presentation_color", $"Presentation color {value} must be a six-digit RGB value.");
        return rgb;
    }

    internal static string NormalizeThemeRgb(string value, out uint? opacity)
    {
        var normalized = value.Trim().TrimStart('#').ToUpperInvariant();
        if (normalized.Length is not (6 or 8) || normalized.Any(character => !Uri.IsHexDigit(character)))
            throw new CodecException("invalid_presentation_color", $"Presentation theme color {value} must be a six- or eight-digit RGB/RGBA value.");
        opacity = normalized.Length == 8
            ? checked((uint)Math.Round(Convert.ToByte(normalized[6..], 16) / 255d * 100_000, MidpointRounding.AwayFromZero))
            : null;
        return normalized[..6];
    }

    internal static string NormalizeScheme(string value)
    {
        var scheme = value.Trim();
        if (!SchemeColors.ContainsKey(scheme))
            throw new CodecException("invalid_presentation_color", $"Presentation scheme color {value} is not a supported theme token.");
        return scheme;
    }

    internal static A.SchemeColorValues SchemeValue(string value) => SchemeColors[NormalizeScheme(value)];

    internal static bool TrySchemeToken(A.SchemeColorValues value, out string token)
    {
        foreach (var entry in SchemeColors)
        {
            if (!entry.Value.Equals(value)) continue;
            token = entry.Key;
            return true;
        }
        token = string.Empty;
        return false;
    }

    internal static bool TrySchemeToken(string value, out string token)
    {
        var candidate = value.Trim();
        foreach (var entry in SchemeColors)
        {
            if (!entry.Key.Equals(candidate, StringComparison.OrdinalIgnoreCase)) continue;
            token = entry.Key;
            return true;
        }
        token = string.Empty;
        return false;
    }
}
