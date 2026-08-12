using System.Security.Cryptography;
using System.Text;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the single p:sld/@show leaf. PresentationML defines absence as shown;
// OfficeKit exposes the inverse Agent-facing state (`hidden`) so callers do
// not need to reason about the native double negative.
internal static class PptxSlideVisibilityCodec
{
    internal sealed record State(bool? Hidden, bool Editable, string SemanticSha256);

    internal static State Read(P.Slide source)
    {
        var attributes = source.GetAttributes()
            .Where(attribute => attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri))
            .ToArray();
        if (attributes.Length == 0) return Known(hidden: false);
        if (attributes.Length != 1) return Opaque(string.Join("\u001f", attributes.Select(attribute => attribute.Value)));
        return (attributes[0].Value ?? string.Empty) switch
        {
            "0" or "false" => Known(hidden: true),
            "1" or "true" => Known(hidden: false),
            var lexical => Opaque(lexical),
        };
    }

    internal static void BuildSourceFree(P.Slide target, PresentationSlide source)
    {
        target.RemoveAttribute("show", string.Empty);
        if (source.HasHidden && source.Hidden) target.Show = false;
    }

    internal static bool ApplySourceBound(P.Slide target, PresentationSlide requested)
    {
        var actual = Read(target);
        if (!actual.Editable || !requested.HasHidden)
            throw new CodecException(
                "unsupported_presentation_slide_visibility_edit",
                "Imported slide visibility is opaque or missing its source-proven semantic and cannot be edited safely.");
        if (requested.Hidden == actual.Hidden) return false;
        target.RemoveAttribute("show", string.Empty);
        if (requested.Hidden) target.Show = false;
        return true;
    }

    internal static bool Matches(PresentationSlide requested, P.Slide source, P.Slide output)
    {
        var original = Read(source);
        var actual = Read(output);
        if (!original.Editable)
            return !requested.HasHidden && !actual.Editable &&
                actual.SemanticSha256.Equals(original.SemanticSha256, StringComparison.OrdinalIgnoreCase);
        return actual.Editable && requested.HasHidden && actual.Hidden == requested.Hidden;
    }

    private static State Known(bool hidden) => new(hidden, true, Hash(hidden ? "hidden" : "visible"));

    private static State Opaque(string lexical) => new(null, false, Hash("opaque\0" + lexical));

    private static string Hash(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
}
