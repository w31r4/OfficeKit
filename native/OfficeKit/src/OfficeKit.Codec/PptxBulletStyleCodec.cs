using System.Globalization;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns the three independent direct DrawingML marker-style choices. Unset wire
// choices preserve unknown source styling. Absent font/color choices clear
// modeled direct declarations; inherited resolution remains host-owned.
internal static class PptxBulletStyleCodec
{
    private const double MaxSizePoints = 768;

    internal static void Read(PresentationTextParagraph target, A.TextParagraphPropertiesType? source)
    {
        if (source is null) return;
        var font = FontChoices(source).ToArray();
        if (font.Length == 1 && ModeledFont(font[0]))
        {
            if (font[0] is A.BulletFont specified) target.BulletFontFamily = specified.Typeface!.Value!;
            else target.BulletFontFollowText = true;
        }

        var color = ColorChoices(source).ToArray();
        if (color.Length == 1 && ModeledColor(color[0]))
        {
            if (color[0] is A.BulletColor specified)
            {
                if (specified.GetFirstChild<A.RgbColorModelHex>() is { } rgb)
                {
                    target.BulletColorRgb = PptxColor.Normalize(rgb.Val!.Value!);
                    if (TryDirectOpacity(rgb, out var opacity) && opacity is { } value)
                        target.BulletColorOpacityThousandthPercent = value;
                }
                else if (specified.GetFirstChild<A.SchemeColor>() is { } scheme && TryScheme(scheme, out var token))
                {
                    target.BulletColorScheme = token;
                    if (TryDirectOpacity(scheme, out var opacity) && opacity is { } value)
                        target.BulletColorOpacityThousandthPercent = value;
                }
            }
            else target.BulletColorFollowText = true;
        }

        var size = SizeChoices(source).ToArray();
        if (size.Length == 1 && ModeledSize(size[0]))
        {
            switch (size[0])
            {
                case A.BulletSizePoints points:
                    TrySizeValue(points, 100, 76_800, out var pointValue);
                    target.BulletSizePoints = pointValue / 100d;
                    break;
                case A.BulletSizePercentage percent:
                    TrySizeValue(percent, 25_000, 400_000, out var percentValue);
                    target.BulletSizePercent = percentValue / 100_000d;
                    break;
                default:
                    target.BulletSizeFollowText = true;
                    break;
            }
        }
    }

    internal static void Validate(PresentationTextParagraph paragraph)
    {
        switch (paragraph.BulletFontCase)
        {
            case PresentationTextParagraph.BulletFontOneofCase.None:
                break;
            case PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily:
                if (!ValidFontFamily(paragraph.BulletFontFamily))
                    throw Invalid("Presentation bullet font family must contain 1 through 255 XML-compatible Unicode scalars and non-whitespace content.");
                break;
            case PresentationTextParagraph.BulletFontOneofCase.BulletFontFollowText:
                if (!paragraph.BulletFontFollowText) throw Invalid("Presentation bullet_font_follow_text must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown bullet-font case.");
        }

        switch (paragraph.BulletColorCase)
        {
            case PresentationTextParagraph.BulletColorOneofCase.None:
                break;
            case PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb:
                _ = PptxColor.Normalize(paragraph.BulletColorRgb);
                break;
            case PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme:
                _ = PptxColor.NormalizeScheme(paragraph.BulletColorScheme);
                break;
            case PresentationTextParagraph.BulletColorOneofCase.BulletColorFollowText:
                if (!paragraph.BulletColorFollowText) throw Invalid("Presentation bullet_color_follow_text must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown bullet-color case.");
        }
        if (paragraph.HasBulletColorOpacityThousandthPercent &&
            (paragraph.BulletColorCase is not (PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb or
             PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme) ||
             paragraph.BulletColorOpacityThousandthPercent > 100_000))
            throw Invalid("Presentation bullet color opacity requires an RGB or theme color and must be from 0 to 100000.");

        switch (paragraph.BulletSizeCase)
        {
            case PresentationTextParagraph.BulletSizeOneofCase.None:
                break;
            case PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints:
                if (!double.IsFinite(paragraph.BulletSizePoints) || paragraph.BulletSizePoints < 1 || paragraph.BulletSizePoints > MaxSizePoints)
                    throw Invalid($"Presentation bullet size must be from 1 through {MaxSizePoints} points.");
                break;
            case PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent:
                if (!double.IsFinite(paragraph.BulletSizePercent) || paragraph.BulletSizePercent < 0.25 || paragraph.BulletSizePercent > 4)
                    throw Invalid("Presentation bullet size percentage must be from 0.25 through 4.");
                break;
            case PresentationTextParagraph.BulletSizeOneofCase.BulletSizeFollowText:
                if (!paragraph.BulletSizeFollowText) throw Invalid("Presentation bullet_size_follow_text must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown bullet-size case.");
        }
    }

    internal static bool HasModeledStyle(PresentationTextParagraph paragraph) =>
        paragraph.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None ||
        paragraph.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None ||
        paragraph.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None;

    internal static void Append(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        if (source.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None) target.AddChild(BuildColor(source), true);
        if (source.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None) target.AddChild(BuildSize(source), true);
        if (source.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None) target.AddChild(BuildFont(source), true);
    }

    internal static void Apply(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        ApplyChoice(target, source, source.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None, ColorChoices, ModeledColor, BuildColor, "color", clearAbsent: true);
        ApplyChoice(target, source, source.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None, SizeChoices, ModeledSize, BuildSize, "size", clearAbsent: true);
        ApplyFont(target, source);
    }

    private static void ApplyFont(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        var existing = FontChoices(target).ToArray();
        var requested = source.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None;
        if (existing.Length > 1 || existing.Any(choice => !ModeledFont(choice)))
        {
            if (requested)
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled or malformed bullet font.");
            return;
        }
        if (!requested)
        {
            foreach (var choice in existing) choice.Remove();
            return;
        }
        if (existing.FirstOrDefault() is A.BulletFont font &&
            source.BulletFontCase == PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily)
        {
            if (font.Typeface!.Value != source.BulletFontFamily) font.Typeface = source.BulletFontFamily;
            return;
        }
        if (existing.FirstOrDefault() is A.BulletFontText &&
            source.BulletFontCase == PresentationTextParagraph.BulletFontOneofCase.BulletFontFollowText) return;
        foreach (var choice in existing) choice.Remove();
        target.AddChild(BuildFont(source), true);
    }

    internal static void Scrub(A.TextParagraphPropertiesType target)
    {
        ScrubChoice(target, ColorChoices, ModeledColor);
        ScrubChoice(target, SizeChoices, ModeledSize);
        ScrubChoice(target, FontChoices, ModeledFont);
    }

    private static void ApplyChoice(
        A.TextParagraphPropertiesType target,
        PresentationTextParagraph source,
        bool requested,
        Func<A.TextParagraphPropertiesType, IEnumerable<OpenXmlElement>> choices,
        Func<OpenXmlElement, bool> modeled,
        Func<PresentationTextParagraph, OpenXmlElement> build,
        string kind,
        bool clearAbsent = false)
    {
        if (!requested)
        {
            if (clearAbsent) ScrubChoice(target, choices, modeled);
            return;
        }
        var existing = choices(target).ToArray();
        if (existing.Length > 1 || existing.Any(choice => !modeled(choice)))
            throw new CodecException("unsupported_presentation_edit", $"Source-preserving PPTX export cannot replace an unmodeled or malformed bullet {kind}.");
        var replacement = build(source);
        if (existing.Length == 1)
        {
            var previous = new PresentationTextParagraph();
            Read(previous, target);
            // Compare modeled meaning without normalizing an unchanged source node.
            if (build(previous).OuterXml == replacement.OuterXml) return;
        }
        foreach (var choice in existing) choice.Remove();
        target.AddChild(replacement, true);
    }

    private static void ScrubChoice(
        A.TextParagraphPropertiesType target,
        Func<A.TextParagraphPropertiesType, IEnumerable<OpenXmlElement>> choices,
        Func<OpenXmlElement, bool> modeled)
    {
        var existing = choices(target).ToArray();
        if (existing.Length == 1 && modeled(existing[0])) existing[0].Remove();
    }

    private static OpenXmlElement BuildFont(PresentationTextParagraph source) => source.BulletFontCase switch
    {
        PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily => new A.BulletFont { Typeface = source.BulletFontFamily },
        PresentationTextParagraph.BulletFontOneofCase.BulletFontFollowText => new A.BulletFontText(),
        _ => throw Invalid("Presentation paragraph has no modeled bullet-font style."),
    };

    private static OpenXmlElement BuildColor(PresentationTextParagraph source) => source.BulletColorCase switch
    {
        PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb => new A.BulletColor(Color(source, new A.RgbColorModelHex { Val = PptxColor.Normalize(source.BulletColorRgb) })),
        PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme => new A.BulletColor(Color(source, new A.SchemeColor { Val = PptxColor.SchemeValue(source.BulletColorScheme) })),
        PresentationTextParagraph.BulletColorOneofCase.BulletColorFollowText => new A.BulletColorText(),
        _ => throw Invalid("Presentation paragraph has no modeled bullet-color style."),
    };

    private static OpenXmlElement Color(PresentationTextParagraph source, OpenXmlElement color)
    {
        if (source.HasBulletColorOpacityThousandthPercent)
            color.Append(new A.Alpha { Val = checked((int)source.BulletColorOpacityThousandthPercent) });
        return color;
    }

    private static OpenXmlElement BuildSize(PresentationTextParagraph source) => source.BulletSizeCase switch
    {
        PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints => new A.BulletSizePoints { Val = checked((int)Math.Round(source.BulletSizePoints * 100)) },
        PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent => new A.BulletSizePercentage { Val = checked((int)Math.Round(source.BulletSizePercent * 100_000)) },
        PresentationTextParagraph.BulletSizeOneofCase.BulletSizeFollowText => new A.BulletSizeText(),
        _ => throw Invalid("Presentation paragraph has no modeled bullet-size style."),
    };

    private static IEnumerable<OpenXmlElement> FontChoices(A.TextParagraphPropertiesType source) =>
        source.ChildElements.Where(child => child is A.BulletFont or A.BulletFontText);

    private static IEnumerable<OpenXmlElement> ColorChoices(A.TextParagraphPropertiesType source) =>
        source.ChildElements.Where(child => child is A.BulletColor or A.BulletColorText);

    private static IEnumerable<OpenXmlElement> SizeChoices(A.TextParagraphPropertiesType source) =>
        source.ChildElements.Where(child => child is A.BulletSizePoints or A.BulletSizePercentage or A.BulletSizeText);

    private static bool ModeledFont(OpenXmlElement source) => source switch
    {
        A.BulletFont font => SimpleAttribute(font, "typeface") && font.GetAttributes()[0].NamespaceUri.Length == 0 && ValidFontFamily(font.Typeface?.Value),
        A.BulletFontText follow => Empty(follow),
        _ => false,
    };

    private static bool ValidFontFamily(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return value.EnumerateRunes().Count() <= 255;
        }
        catch (XmlException) { return false; }
    }

    private static bool ModeledColor(OpenXmlElement source) => source switch
    {
        A.BulletColor color when EmptyAttributes(color) && color.ChildElements.Count == 1 && color.GetFirstChild<A.RgbColorModelHex>() is { } rgb =>
            ColorAttribute(rgb) && ValidRgb(rgb.Val?.Value) && TryDirectOpacity(rgb, out _),
        A.BulletColor color when EmptyAttributes(color) && color.ChildElements.Count == 1 && color.GetFirstChild<A.SchemeColor>() is { } scheme =>
            TryScheme(scheme, out _) && TryDirectOpacity(scheme, out _),
        A.BulletColorText follow => Empty(follow),
        _ => false,
    };

    private static bool ModeledSize(OpenXmlElement source) => source switch
    {
        A.BulletSizePoints points => TrySizeValue(points, 100, 76_800, out _),
        A.BulletSizePercentage percent => TrySizeValue(percent, 25_000, 400_000, out _),
        A.BulletSizeText follow => Empty(follow),
        _ => false,
    };

    private static bool TrySizeValue(OpenXmlElement source, int minimum, int maximum, out int value)
    {
        value = 0;
        return source.ChildElements.Count == 0 && ColorAttribute(source) &&
            int.TryParse(source.GetAttributes()[0].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) &&
            value >= minimum && value <= maximum;
    }

    private static bool Empty(OpenXmlElement source) => EmptyAttributes(source) && source.ChildElements.Count == 0;

    private static bool EmptyAttributes(OpenXmlElement source) => source.GetAttributes().Count == 0;

    private static bool SimpleAttribute(OpenXmlElement source, string name)
    {
        var attributes = source.GetAttributes();
        return source.ChildElements.Count == 0 && attributes.Count == 1 && attributes[0].LocalName == name;
    }

    private static bool ColorAttribute(OpenXmlElement source)
    {
        var attributes = source.GetAttributes();
        return attributes.Count == 1 && attributes[0].LocalName == "val" && attributes[0].NamespaceUri.Length == 0;
    }

    private static bool TryDirectOpacity(OpenXmlElement source, out uint? opacity)
    {
        opacity = null;
        if (source.ChildElements.Count == 0) return true;
        if (source.ChildElements.Count != 1 || source.FirstChild is not A.Alpha alpha ||
            alpha.ChildElements.Count != 0 || !ColorAttribute(alpha) ||
            !int.TryParse(alpha.GetAttributes()[0].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ||
            value is < 0 or > 100_000) return false;
        opacity = checked((uint)value);
        return true;
    }

    private static bool ValidRgb(string? value) => value is { Length: 6 } && value.All(Uri.IsHexDigit);

    private static bool TryScheme(A.SchemeColor source, out string token)
    {
        token = string.Empty;
        if (!ColorAttribute(source)) return false;
        var raw = source.GetAttributes()[0].Value;
        return PptxColor.TrySchemeToken(raw, out token) &&
            new EnumValue<A.SchemeColorValues>(PptxColor.SchemeValue(token)).InnerText == raw;
    }

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
}
