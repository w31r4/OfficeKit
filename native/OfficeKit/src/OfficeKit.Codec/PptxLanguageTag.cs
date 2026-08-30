using System.Text.RegularExpressions;

namespace OfficeKit.Codec;

internal static partial class PptxLanguageTag
{
    [GeneratedRegex("^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*$", RegexOptions.CultureInvariant)]
    private static partial Regex TokenPattern();

    internal static bool IsValid(string? value) =>
        value is { Length: >= 2 and <= 63 } &&
        !char.IsWhiteSpace(value[0]) &&
        !char.IsWhiteSpace(value[^1]) &&
        TokenPattern().IsMatch(value);

    internal static string Validate(string? value)
    {
        if (!IsValid(value))
            throw new CodecException("invalid_presentation_text", "Presentation text language must be a bounded BCP-47 language tag.");
        return value!;
    }
}
