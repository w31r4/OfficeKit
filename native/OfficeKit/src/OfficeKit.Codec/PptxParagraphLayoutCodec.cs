using System.Globalization;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns direct paragraph coordinates whose absence means inheritance. Explicit
// no_* wire choices carry deletion intent across source-preserving edits.
internal static class PptxParagraphLayoutCodec
{
    private const long MaxCoordinateEmu = 51_206_400;

    internal static void Read(PresentationTextParagraph target, A.TextParagraphPropertiesType? source)
    {
        if (TryMargin(source, out var margin)) target.MarginLeftEmu = margin;
        if (source?.Indent?.Value is { } indent && ValidIndent(indent)) target.IndentEmu = indent;
    }

    internal static bool Supports(A.TextParagraphPropertiesType? source) =>
        source is null ||
        (source.Indent is null || ValidIndent(source.Indent.Value));

    internal static void Validate(PresentationTextParagraph source)
    {
        switch (source.LeftMarginCase)
        {
            case PresentationTextParagraph.LeftMarginOneofCase.None:
                break;
            case PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu:
                if (!ValidMargin(source.MarginLeftEmu)) throw Invalid($"Presentation paragraph left margin must be from 0 through {MaxCoordinateEmu} EMUs.");
                break;
            case PresentationTextParagraph.LeftMarginOneofCase.NoMarginLeft:
                if (!source.NoMarginLeft) throw Invalid("Presentation no_margin_left must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown left-margin case.");
        }

        switch (source.IndentationCase)
        {
            case PresentationTextParagraph.IndentationOneofCase.None:
                break;
            case PresentationTextParagraph.IndentationOneofCase.IndentEmu:
                if (!ValidIndent(source.IndentEmu)) throw Invalid($"Presentation paragraph indent must be from {-MaxCoordinateEmu} through {MaxCoordinateEmu} EMUs.");
                break;
            case PresentationTextParagraph.IndentationOneofCase.NoIndent:
                if (!source.NoIndent) throw Invalid("Presentation no_indent must be true when selected.");
                break;
            default:
                throw Invalid("Presentation paragraph contains an unknown indentation case.");
        }
    }

    internal static bool HasAuthoredLayout(PresentationTextParagraph source) =>
        source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu ||
        source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu;

    internal static void Append(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        if (source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
            target.LeftMargin = checked((int)source.MarginLeftEmu);
        if (source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
            target.Indent = checked((int)source.IndentEmu);
    }

    internal static void Apply(A.TextParagraphPropertiesType target, PresentationTextParagraph source)
    {
        if (source.LeftMarginCase != PresentationTextParagraph.LeftMarginOneofCase.None)
        {
            var modeled = TryMargin(target, out var margin);
            if (target.LeftMargin is not null && !modeled) throw Unsupported("left margin");
            if (source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
            {
                if (!modeled || margin != source.MarginLeftEmu) target.LeftMargin = checked((int)source.MarginLeftEmu);
            }
            else target.LeftMargin = null;
        }
        if (source.IndentationCase != PresentationTextParagraph.IndentationOneofCase.None)
        {
            if (target.Indent is not null && !ValidIndent(target.Indent.Value)) throw Unsupported("indent");
            if (source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
            {
                if (target.Indent?.Value != source.IndentEmu) target.Indent = checked((int)source.IndentEmu);
            }
            else target.Indent = null;
        }
    }

    internal static void Scrub(A.TextParagraphPropertiesType target)
    {
        if (TryMargin(target, out _)) target.LeftMargin = null;
        if (target.Indent?.Value is { } indent && ValidIndent(indent)) target.Indent = null;
    }

    private static bool TryMargin(A.TextParagraphPropertiesType? source, out long value) =>
        long.TryParse(source?.GetAttributes().FirstOrDefault(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "marL").Value,
            NumberStyles.Integer, CultureInfo.InvariantCulture, out value) && ValidMargin(value);

    private static bool ValidMargin(long value) => value is >= 0 and <= MaxCoordinateEmu;
    private static bool ValidIndent(long value) => value is >= -MaxCoordinateEmu and <= MaxCoordinateEmu;
    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
    private static CodecException Unsupported(string kind) => new("unsupported_presentation_edit", $"Source-preserving PPTX export cannot replace an unmodeled paragraph {kind}.");
}
