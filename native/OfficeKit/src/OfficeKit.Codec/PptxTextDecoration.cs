using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Direct DrawingML text decorations are intentionally a small, source-bound
// slice.  Underline effects (for example uFill/uLn) carry their own graph and
// stay opaque; only a plain token on a:rPr/defRPr can be edited in place.
internal static class PptxTextDecoration
{
    private const int MaxKerningHundredthsPoints = 76_800;
    private const int MinSpacingHundredthsPoints = -76_800;
    private const int MaxSpacingHundredthsPoints = 76_800;
    private const int MinBaselineThousandthsPercent = -400_000;
    private const int MaxBaselineThousandthsPercent = 400_000;
    private static readonly HashSet<string> UnderlineValues = new(StringComparer.Ordinal)
    {
        "none", "words", "sng", "dbl", "heavy", "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong", "dashLongHeavy",
        "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy", "wavy", "wavyHeavy", "wavyDbl",
    };

    private static readonly HashSet<string> StrikeValues = new(StringComparer.Ordinal)
    {
        "noStrike", "sngStrike", "dblStrike",
    };

    private static readonly HashSet<string> CapsValues = new(StringComparer.Ordinal)
    {
        "none", "small", "all",
    };

    internal static string NormalizeUnderline(string value)
    {
        var token = value?.Trim() ?? string.Empty;
        if (!UnderlineValues.Contains(token))
            throw new CodecException("invalid_presentation_text", $"Unsupported Presentation underline token {value}.");
        return token;
    }

    internal static string NormalizeStrike(string value)
    {
        var token = value?.Trim() ?? string.Empty;
        if (!StrikeValues.Contains(token))
            throw new CodecException("invalid_presentation_text", $"Unsupported Presentation strike token {value}.");
        return token;
    }

    internal static bool TryUnderline(A.TextCharacterPropertiesType? source, out string value)
    {
        var raw = source?.Underline?.ToString();
        if (raw is not null && UnderlineValues.Contains(raw) && source is not null && !HasUnderlineEffects(source))
        {
            value = raw;
            return true;
        }
        value = string.Empty;
        return false;
    }

    internal static bool TryStrike(A.TextCharacterPropertiesType? source, out string value)
    {
        var raw = source?.Strike?.ToString();
        if (raw is not null && StrikeValues.Contains(raw))
        {
            value = raw;
            return true;
        }
        value = string.Empty;
        return false;
    }

    internal static bool IsUnderlineToken(string value) => UnderlineValues.Contains(value);

    internal static bool IsStrikeToken(string value) => StrikeValues.Contains(value);

    internal static bool TryCaps(A.TextCharacterPropertiesType? source, out string value)
    {
        var raw = source?.Capital?.InnerText;
        if (raw is not null && CapsValues.Contains(raw))
        {
            value = raw;
            return true;
        }
        value = string.Empty;
        return false;
    }

    internal static string NormalizeCaps(string value)
    {
        var token = value?.Trim() ?? string.Empty;
        if (!CapsValues.Contains(token))
            throw new CodecException("invalid_presentation_text", $"Unsupported Presentation capitalization token {value}.");
        return token;
    }

    internal static bool IsCapsToken(string value) => CapsValues.Contains(value);

    internal static bool TryKerning(A.TextCharacterPropertiesType? source, out string value)
    {
        if (source?.Kerning?.Value is { } raw && raw >= 0 && raw <= MaxKerningHundredthsPoints)
        {
            value = raw.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return true;
        }
        value = string.Empty;
        return false;
    }

    internal static string NormalizeKerning(string value)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var raw) ||
            raw < 0 || raw > MaxKerningHundredthsPoints ||
            raw.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_text", $"Unsupported Presentation kerning token {value}.");
        return raw.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    internal static bool IsKerningToken(string value)
    {
        try
        {
            _ = NormalizeKerning(value);
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    internal static bool TrySpacing(A.TextCharacterPropertiesType? source, out string value)
    {
        if (source?.Spacing?.Value is { } raw && raw >= MinSpacingHundredthsPoints && raw <= MaxSpacingHundredthsPoints)
        {
            value = raw.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return true;
        }
        value = string.Empty;
        return false;
    }

    internal static string NormalizeSpacing(string value)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var raw) ||
            raw < MinSpacingHundredthsPoints || raw > MaxSpacingHundredthsPoints ||
            raw.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_text", $"Unsupported Presentation character spacing token {value}.");
        return raw.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    internal static bool IsSpacingToken(string value)
    {
        try
        {
            _ = NormalizeSpacing(value);
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    internal static bool TryBaseline(A.TextCharacterPropertiesType? source, out string value)
    {
        if (source?.Baseline?.Value is { } raw && raw >= MinBaselineThousandthsPercent && raw <= MaxBaselineThousandthsPercent)
        {
            value = raw.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return true;
        }
        value = string.Empty;
        return false;
    }

    internal static string NormalizeBaseline(string value)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var raw) ||
            raw < MinBaselineThousandthsPercent || raw > MaxBaselineThousandthsPercent ||
            raw.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_text", $"Unsupported Presentation baseline token {value}.");
        return raw.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    internal static bool IsBaselineToken(string value)
    {
        try
        {
            _ = NormalizeBaseline(value);
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static bool HasUnderlineEffects(A.TextCharacterPropertiesType source) =>
        source.ChildElements.Any(child => child.LocalName is "uFillTx" or "uFill" or "uLnTx" or "uLn");
}
