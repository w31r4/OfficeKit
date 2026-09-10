using System.Globalization;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns the direct DrawingML list-marker choice. Picture bytes are durable
// assets while each slide context owns its package relationships.
internal static class PptxBulletCodec
{
    private static readonly HashSet<string> AutoNumberSchemes = new(StringComparer.Ordinal)
    {
        "alphaLcParenBoth", "alphaLcParenR", "alphaLcPeriod", "alphaUcParenBoth", "alphaUcParenR", "alphaUcPeriod",
        "arabic1Minus", "arabic2Minus", "arabicDbPeriod", "arabicDbPlain", "arabicParenBoth", "arabicParenR", "arabicPeriod", "arabicPlain",
        "circleNumDbPlain", "circleNumWdBlackPlain", "circleNumWdWhitePlain", "ea1ChsPeriod", "ea1ChsPlain", "ea1ChtPeriod", "ea1ChtPlain",
        "ea1JpnChsDbPeriod", "ea1JpnKorPeriod", "ea1JpnKorPlain", "hebrew2Minus", "hindiAlpha1Period", "hindiAlphaPeriod",
        "hindiNumParenR", "hindiNumPeriod", "romanLcParenBoth", "romanLcParenR", "romanLcPeriod", "romanUcParenBoth", "romanUcParenR", "romanUcPeriod",
        "thaiAlphaParenBoth", "thaiAlphaParenR", "thaiAlphaPeriod", "thaiNumParenBoth", "thaiNumParenR", "thaiNumPeriod",
    };

    internal static void Read(PresentationTextParagraph target, A.TextParagraphPropertiesType? source, PptxPartContext? context)
    {
        if (source is null) return;
        var choices = BulletChoices(source).ToArray();
        if (choices.Length != 1 || !Modeled(choices[0], context)) return;
        switch (choices[0])
        {
            case A.NoBullet:
                target.NoBullet = true;
                break;
            case A.CharacterBullet character:
                target.BulletCharacter = character.Char!.Value!;
                break;
            case A.AutoNumberedBullet autoNumber:
                target.AutoNumber = new PresentationAutoNumberBullet { Scheme = Scheme(autoNumber) };
                if (TryStartAt(autoNumber, out var startAt) && startAt is { } value) target.AutoNumber.StartAt = value;
                break;
            case A.PictureBullet picture when context!.TryReadPicture(picture, out var modeled):
                target.PictureBullet = modeled;
                break;
        }
    }

    internal static void Validate(PresentationTextParagraph paragraph)
    {
        switch (paragraph.BulletCase)
        {
            case PresentationTextParagraph.BulletOneofCase.None:
                return;
            case PresentationTextParagraph.BulletOneofCase.NoBullet:
                if (!paragraph.NoBullet) throw Invalid("Presentation no_bullet must be true when selected.");
                return;
            case PresentationTextParagraph.BulletOneofCase.BulletCharacter:
                if (!ValidCharacter(paragraph.BulletCharacter)) throw Invalid("Presentation bullet character must contain exactly one XML-compatible Unicode scalar value.");
                return;
            case PresentationTextParagraph.BulletOneofCase.AutoNumber:
                if (paragraph.AutoNumber is null || !AutoNumberSchemes.Contains(paragraph.AutoNumber.Scheme))
                    throw Invalid($"Unsupported Presentation auto-number scheme {paragraph.AutoNumber?.Scheme ?? "(missing)"}.");
                if (paragraph.AutoNumber.HasStartAt && (paragraph.AutoNumber.StartAt < 1 || paragraph.AutoNumber.StartAt > 32_767))
                    throw Invalid("Presentation auto-number start_at must be from 1 through 32767.");
                return;
            case PresentationTextParagraph.BulletOneofCase.PictureBullet:
                PptxPartContext.ValidatePicture(paragraph.PictureBullet);
                return;
            default:
                throw Invalid("Presentation paragraph contains an unknown bullet case.");
        }
    }

    internal static bool HasModeledBullet(PresentationTextParagraph paragraph) =>
        paragraph.BulletCase != PresentationTextParagraph.BulletOneofCase.None;

    internal static bool IsAutoNumberScheme(string value) => AutoNumberSchemes.Contains(value);

    internal static void Append(A.TextParagraphPropertiesType target, PresentationTextParagraph source, PptxPartContext? context)
    {
        if (!HasModeledBullet(source)) return;
        target.AddChild(Build(source, context), true);
    }

    internal static void Apply(A.TextParagraphPropertiesType target, PresentationTextParagraph source, PptxPartContext context)
    {
        if (!HasModeledBullet(source)) return;
        var existing = BulletChoices(target).ToArray();
        if (existing.Length > 1 || existing.Any(choice => !Modeled(choice, context)))
            throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot replace an unmodeled or malformed list marker.");
        if (existing.Length == 1 && existing[0] is A.CharacterBullet character &&
            source.BulletCase == PresentationTextParagraph.BulletOneofCase.BulletCharacter)
        {
            if (character.Char!.Value != source.BulletCharacter) character.Char = source.BulletCharacter;
            return;
        }
        if (existing.Length == 1 && existing[0] is A.AutoNumberedBullet autoNumber &&
            source.BulletCase == PresentationTextParagraph.BulletOneofCase.AutoNumber)
        {
            // Number format and optional start each own one attribute, not
            // the marker element. Retain unknown XML and unchanged spelling.
            if (Scheme(autoNumber) != source.AutoNumber.Scheme)
                autoNumber.Type = new A.TextAutoNumberSchemeValues(source.AutoNumber.Scheme);
            TryStartAt(autoNumber, out var currentStartAt);
            if (source.AutoNumber.HasStartAt)
            {
                if (currentStartAt != source.AutoNumber.StartAt)
                    autoNumber.StartAt = checked((int)source.AutoNumber.StartAt);
            }
            else if (currentStartAt is not null) autoNumber.StartAt = null;
            return;
        }
        if (existing.Length == 1 && existing[0] is A.PictureBullet picture &&
            source.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
            context.TryReadPicture(picture, out var current) && current.Equals(source.PictureBullet)) return;
        foreach (var choice in existing) choice.Remove();
        target.AddChild(Build(source, context), true);
    }

    internal static void Scrub(A.TextParagraphPropertiesType target, PptxPartContext? context)
    {
        var choices = BulletChoices(target).ToArray();
        if (choices.Length == 1 && (Modeled(choices[0], context) ||
            choices[0] is A.PictureBullet picture && context?.IsPictureRelationship(picture) == true))
            choices[0].Remove();
    }

    private static OpenXmlElement Build(PresentationTextParagraph source, PptxPartContext? context) => source.BulletCase switch
    {
        PresentationTextParagraph.BulletOneofCase.NoBullet => new A.NoBullet(),
        PresentationTextParagraph.BulletOneofCase.BulletCharacter => new A.CharacterBullet { Char = source.BulletCharacter },
        PresentationTextParagraph.BulletOneofCase.AutoNumber => new A.AutoNumberedBullet
        {
            Type = new A.TextAutoNumberSchemeValues(source.AutoNumber.Scheme),
            StartAt = source.AutoNumber.HasStartAt ? checked((int)source.AutoNumber.StartAt) : null,
        },
        PresentationTextParagraph.BulletOneofCase.PictureBullet => context?.BuildPicture(source.PictureBullet) ??
            throw Invalid("Presentation picture bullet authoring requires a slide context."),
        _ => throw Invalid("Presentation paragraph has no modeled list marker."),
    };

    private static IEnumerable<OpenXmlElement> BulletChoices(A.TextParagraphPropertiesType source) =>
        source.ChildElements.Where(child => child is A.NoBullet or A.CharacterBullet or A.AutoNumberedBullet or A.PictureBullet);

    private static bool Modeled(OpenXmlElement source, PptxPartContext? context) => source switch
    {
        A.NoBullet => true,
        A.CharacterBullet character => ValidCharacter(character.Char?.Value),
        A.AutoNumberedBullet autoNumber => AutoNumberSchemes.Contains(Scheme(autoNumber)) && TryStartAt(autoNumber, out _),
        A.PictureBullet picture => context is not null && context.TryReadPicture(picture, out _),
        _ => false,
    };

    private static bool ValidCharacter(string? value)
    {
        if (string.IsNullOrEmpty(value) || !Rune.TryGetRuneAt(value, 0, out var rune) ||
            rune.Utf16SequenceLength != value.Length) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return true;
        }
        catch (XmlException) { return false; }
    }

    private static string Scheme(A.AutoNumberedBullet source) => source.Type?.InnerText ?? string.Empty;

    private static bool TryStartAt(A.AutoNumberedBullet source, out uint? startAt)
    {
        startAt = null;
        var token = source.GetAttributes().FirstOrDefault(attribute =>
            attribute.NamespaceUri.Length == 0 && attribute.LocalName == "startAt").Value;
        if (token is null) return true;
        if (!uint.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) || value is < 1 or > 32_767)
            return false;
        startAt = value;
        return true;
    }

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
}
