using System.Globalization;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Coordinates the modeled CT_TextParagraphProperties subset shared by direct
// a:pPr paragraphs and a:lstStyle level defaults. Child-specific codecs retain
// ownership of their native elements and source-preserving edit rules.
internal static class PptxParagraphPropertiesCodec
{
    internal static void Read(
        PresentationTextParagraph target,
        A.TextParagraphPropertiesType? source,
        PptxPartContext? slideContext,
        bool readLevel)
    {
        if (FontAlignmentName(source) is { Length: > 0 } fontAlignment) target.FontAlignment = fontAlignment;
        if (RightToLeft(source) is { } rightToLeft) target.RightToLeft = rightToLeft;
        if (HangingPunctuation(source) is { } hangingPunctuation) target.HangingPunctuation = hangingPunctuation;
        if (LatinLineBreak(source) is { } latinLineBreak) target.LatinLineBreak = latinLineBreak;
        if (EastAsianLineBreak(source) is { } eastAsianLineBreak) target.EastAsianLineBreak = eastAsianLineBreak;
        if (readLevel && TryLevel(source, out var level)) target.Level = level;
        if (AlignmentName(source) is { Length: > 0 } name)
            target.Alignment = name;
        PptxParagraphLayoutCodec.Read(target, source);
        PptxParagraphSpacingCodec.Read(target, source);
        PptxBulletCodec.Read(target, source, slideContext);
        PptxBulletStyleCodec.Read(target, source);
        PptxTextCodec.ReadTabStops(target, source);
        PptxDefaultRunStyleCodec.Read(target, source);
    }

    internal static bool Supports(A.TextParagraphPropertiesType? source) =>
        PptxDefaultRunStyleCodec.Supports(source);

    internal static void Validate(PresentationTextParagraph source, bool requireLevel)
    {
        if (requireLevel && !source.HasLevel)
            throw Invalid("Presentation list style must identify a level from 0 through 8.");
        if (source.HasLevel && source.Level > 8)
            throw Invalid("Presentation paragraph level must be from 0 through 8.");
        if (source.HasFontAlignment) _ = ParseFontAlignment(source.FontAlignment);
        if (source.HasAlignment) _ = ParseAlignment(source.Alignment);
        PptxParagraphLayoutCodec.Validate(source);
        PptxParagraphSpacingCodec.Validate(source);
        PptxDefaultRunStyleCodec.Validate(source);
        PptxBulletCodec.Validate(source);
        PptxBulletStyleCodec.Validate(source);
        PptxTextCodec.ValidateTabStops(source);
    }

    internal static bool HasAuthoredProperties(PresentationTextParagraph source, bool includeLevel) =>
        includeLevel && source.HasLevel ||
        source.HasAlignment || source.HasRightToLeft || source.HasFontAlignment || source.HasHangingPunctuation || source.HasLatinLineBreak || source.HasEastAsianLineBreak ||
        PptxParagraphLayoutCodec.HasAuthoredLayout(source) ||
        PptxParagraphSpacingCodec.HasAuthoredSpacing(source) ||
        PptxBulletCodec.HasModeledBullet(source) ||
        PptxBulletStyleCodec.HasModeledStyle(source) ||
        PptxDefaultRunStyleCodec.HasAuthoredStyle(source) ||
        source.TabStops.Count > 0;

    internal static bool HasModeledProperties(PresentationTextParagraph source) =>
        source.HasAlignment || source.HasRightToLeft || source.HasFontAlignment || source.HasHangingPunctuation || source.HasLatinLineBreak || source.HasEastAsianLineBreak ||
        source.LeftMarginCase != PresentationTextParagraph.LeftMarginOneofCase.None ||
        source.RightMarginCase != PresentationTextParagraph.RightMarginOneofCase.None ||
        source.IndentationCase != PresentationTextParagraph.IndentationOneofCase.None ||
        source.LineSpacingCase != PresentationTextParagraph.LineSpacingOneofCase.None ||
        source.SpaceBeforeCase != PresentationTextParagraph.SpaceBeforeOneofCase.None ||
        source.SpaceAfterCase != PresentationTextParagraph.SpaceAfterOneofCase.None ||
        source.BulletCase != PresentationTextParagraph.BulletOneofCase.None ||
        source.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None ||
        source.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None ||
        source.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None ||
        source.TabStops.Count > 0 || source.HasNoTabStops ||
        source.DefaultRunStyleCase != PresentationTextParagraph.DefaultRunStyleOneofCase.None;

    internal static void Append(
        A.TextParagraphPropertiesType target,
        PresentationTextParagraph source,
        PptxPartContext? slideContext,
        bool includeLevel)
    {
        if (source.HasFontAlignment) target.FontAlignment = ParseFontAlignment(source.FontAlignment);
        if (source.HasRightToLeft) target.RightToLeft = source.RightToLeft;
        if (source.HasHangingPunctuation) target.Height = source.HangingPunctuation;
        if (source.HasLatinLineBreak) target.LatinLineBreak = source.LatinLineBreak;
        if (source.HasEastAsianLineBreak) target.EastAsianLineBreak = source.EastAsianLineBreak;
        if (includeLevel && source.HasLevel) target.Level = checked((int)source.Level);
        if (source.HasAlignment) target.Alignment = ParseAlignment(source.Alignment);
        PptxParagraphLayoutCodec.Append(target, source);
        PptxParagraphSpacingCodec.Append(target, source);
        PptxBulletStyleCodec.Append(target, source);
        PptxBulletCodec.Append(target, source, slideContext);
        PptxTextCodec.AppendTabStops(target, source);
        PptxDefaultRunStyleCodec.Append(target, source);
    }

    internal static void Apply(
        A.TextParagraphPropertiesType target,
        PresentationTextParagraph source,
        PptxPartContext slideContext,
        bool includeLevel)
    {
        if (includeLevel)
        {
            var modeled = TryLevel(target, out var level);
            if (source.HasLevel)
            {
                if (!modeled && target.GetAttributes().Any(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "lvl"))
                    throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled paragraph level.");
                if (!modeled || level != source.Level) target.Level = checked((int)source.Level);
            }
            else if (modeled) target.Level = null;
        }
        var fontAlignment = FontAlignmentName(target);
        if (source.HasFontAlignment)
        {
            if (fontAlignment.Length == 0 && target.GetAttributes().Any(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "fontAlgn"))
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled paragraph font alignment.");
            if (fontAlignment != source.FontAlignment) target.FontAlignment = ParseFontAlignment(source.FontAlignment);
        }
        else if (fontAlignment.Length > 0) target.FontAlignment = null;
        var rightToLeft = RightToLeft(target);
        if (source.HasRightToLeft)
        {
            if (rightToLeft is null && target.GetAttributes().Any(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "rtl"))
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled paragraph direction.");
            if (rightToLeft != source.RightToLeft) target.RightToLeft = source.RightToLeft;
        }
        else if (rightToLeft is not null) target.RightToLeft = null;
        var hangingPunctuation = HangingPunctuation(target);
        if (source.HasHangingPunctuation)
        {
            if (hangingPunctuation is null && target.GetAttributes().Any(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "hangingPunct"))
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled paragraph hanging punctuation.");
            if (hangingPunctuation != source.HangingPunctuation) target.Height = source.HangingPunctuation;
        }
        else if (hangingPunctuation is not null) target.Height = null;
        var latinLineBreak = LatinLineBreak(target);
        if (source.HasLatinLineBreak)
        {
            if (latinLineBreak is null && target.GetAttributes().Any(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "latinLnBrk"))
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled paragraph Latin line break.");
            if (latinLineBreak != source.LatinLineBreak) target.LatinLineBreak = source.LatinLineBreak;
        }
        else if (latinLineBreak is not null) target.LatinLineBreak = null;
        var eastAsianLineBreak = EastAsianLineBreak(target);
        if (source.HasEastAsianLineBreak)
        {
            if (eastAsianLineBreak is null && target.GetAttributes().Any(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "eaLnBrk"))
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled paragraph East Asian line break.");
            if (eastAsianLineBreak != source.EastAsianLineBreak) target.EastAsianLineBreak = source.EastAsianLineBreak;
        }
        else if (eastAsianLineBreak is not null) target.EastAsianLineBreak = null;
        var alignment = AlignmentName(target);
        if (source.HasAlignment)
        {
            if (alignment.Length == 0 && target.GetAttributes().Any(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "algn"))
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled paragraph alignment.");
            if (alignment != source.Alignment) target.Alignment = ParseAlignment(source.Alignment);
        }
        else if (alignment.Length > 0) target.Alignment = null;
        PptxParagraphLayoutCodec.Apply(target, source);
        PptxParagraphSpacingCodec.Apply(target, source);
        PptxBulletStyleCodec.Apply(target, source);
        PptxBulletCodec.Apply(target, source, slideContext);
        PptxTextCodec.ApplyTabStops(target, source);
        PptxDefaultRunStyleCodec.Apply(target, source);
    }

    internal static void Scrub(A.TextParagraphPropertiesType target, PptxPartContext? slideContext, bool includeLevel)
    {
        if (includeLevel && TryLevel(target, out _)) target.Level = null;
        if (AlignmentName(target).Length > 0) target.Alignment = null;
        if (RightToLeft(target) is not null) target.RightToLeft = null;
        if (HangingPunctuation(target) is not null) target.Height = null;
        if (LatinLineBreak(target) is not null) target.LatinLineBreak = null;
        if (EastAsianLineBreak(target) is not null) target.EastAsianLineBreak = null;
        if (FontAlignmentName(target).Length > 0) target.FontAlignment = null;
        PptxParagraphLayoutCodec.Scrub(target);
        PptxParagraphSpacingCodec.Scrub(target);
        PptxDefaultRunStyleCodec.Scrub(target);
        PptxBulletCodec.Scrub(target, slideContext);
        PptxBulletStyleCodec.Scrub(target);
        if (PptxTextCodec.SupportsTabStops(target)) target.GetFirstChild<A.TabStopList>()?.Remove();
    }

    private static string FontAlignmentName(A.TextParagraphPropertiesType? source) =>
        source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "fontAlgn").Value switch
        {
            "auto" => "auto",
            "t" => "top",
            "ctr" => "center",
            "base" => "baseline",
            "b" => "bottom",
            _ => string.Empty,
        };

    private static A.TextFontAlignmentValues ParseFontAlignment(string value) => value switch
    {
        "auto" => A.TextFontAlignmentValues.Automatic,
        "top" => A.TextFontAlignmentValues.Top,
        "center" => A.TextFontAlignmentValues.Center,
        "baseline" => A.TextFontAlignmentValues.Baseline,
        "bottom" => A.TextFontAlignmentValues.Bottom,
        _ => throw Invalid($"Unsupported Presentation paragraph font alignment {value}."),
    };

    private static bool? RightToLeft(A.TextParagraphPropertiesType? source) =>
        source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "rtl").Value switch
        {
            "0" or "false" => false,
            "1" or "true" => true,
            _ => null,
        };

    // The SDK names the hangingPunct attribute Height; it is a boolean setting,
    // independent of text/shape height and the paragraph's hanging indent.
    private static bool? HangingPunctuation(A.TextParagraphPropertiesType? source) =>
        source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "hangingPunct").Value switch
        {
            "0" or "false" => false,
            "1" or "true" => true,
            _ => null,
        };

    private static bool? LatinLineBreak(A.TextParagraphPropertiesType? source) =>
        source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "latinLnBrk").Value switch
        {
            "0" or "false" => false,
            "1" or "true" => true,
            _ => null,
        };

    private static bool? EastAsianLineBreak(A.TextParagraphPropertiesType? source) =>
        source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "eaLnBrk").Value switch
        {
            "0" or "false" => false,
            "1" or "true" => true,
            _ => null,
        };

    private static bool TryLevel(A.TextParagraphPropertiesType? source, out uint level) =>
        uint.TryParse(source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "lvl").Value,
            NumberStyles.Integer, CultureInfo.InvariantCulture, out level) && level <= 8;

    // Classify raw tokens before accessing SDK enum values: malformed values
    // must remain source-owned instead of aborting projection or unrelated edits.
    private static string AlignmentName(A.TextParagraphPropertiesType? source) =>
        source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "algn").Value switch
        {
            "l" => "left",
            "ctr" => "center",
            "r" => "right",
            "just" => "justify",
            "justLow" => "justifyLow",
            "dist" => "distributed",
            "thaiDist" => "thaiDistributed",
            _ => string.Empty,
        };

    private static A.TextAlignmentTypeValues ParseAlignment(string value) => value switch
    {
        "left" => A.TextAlignmentTypeValues.Left,
        "center" => A.TextAlignmentTypeValues.Center,
        "right" => A.TextAlignmentTypeValues.Right,
        "justify" => A.TextAlignmentTypeValues.Justified,
        "justifyLow" => A.TextAlignmentTypeValues.JustifiedLow,
        "distributed" => A.TextAlignmentTypeValues.Distributed,
        "thaiDistributed" => A.TextAlignmentTypeValues.ThaiDistributed,
        _ => throw Invalid($"Unsupported Presentation paragraph alignment {value}."),
    };

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
}
