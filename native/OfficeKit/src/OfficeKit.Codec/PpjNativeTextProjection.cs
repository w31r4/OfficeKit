using System.Xml;
using System.Xml.Linq;

namespace OfficeKit.Codec;

/// <summary>
/// Describes the same deliberately narrow opaque DrawingML text profile that
/// PptxNativeTextLeafCodec re-proves before token splicing. PPJ receives only
/// ordered decoded strings; the XML and package locator stay native.
/// </summary>
internal static class PpjNativeTextProjection
{
    private const int MaxLeaves = 4_096;
    private const int MaxLeafLength = 32_767;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    internal static bool TryRead(string rawXml, out IReadOnlyList<string> leaves)
    {
        leaves = [];
        if (string.IsNullOrWhiteSpace(rawXml)) return false;
        XElement root;
        try
        {
            using var text = new StringReader(rawXml);
            using var reader = XmlReader.Create(text, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 16 * 1024 * 1024,
            });
            root = XElement.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception) when (exception is XmlException or InvalidOperationException)
        {
            return false;
        }

        var texts = root.Descendants(A + "t").ToArray();
        if (texts.Length is < 1 or > MaxLeaves || texts.Any(text => !ValidText(text.Value))) return false;
        if (root.Name == P + "graphicFrame")
        {
            var tables = root.Descendants(A + "tbl").ToArray();
            if (tables.Length != 1 || tables[0].Descendants(A + "tc").Any() == false ||
                texts.Any(text => text.Parent?.Name != A + "r" || text.Ancestors(A + "tc").Count() != 1))
                return false;
        }
        else if (root.Name == P + "grpSp")
        {
            if (texts.Any(text => text.Parent?.Name != A + "r" ||
                                  text.Ancestors(P + "sp").Count() != 1 ||
                                  text.Ancestors(A + "tc").Any()))
                return false;
        }
        else
        {
            return false;
        }

        leaves = texts.Select(text => text.Value).ToArray();
        return true;
    }

    private static bool ValidText(string value)
    {
        if (value.Length > MaxLeafLength) return false;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (character <= '\b' || character is '\v' or '\f' || character is >= '\u000E' and <= '\u001F') return false;
            if (char.IsHighSurrogate(character))
            {
                if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1])) return false;
                index++;
            }
            else if (char.IsLowSurrogate(character))
            {
                return false;
            }
        }
        return true;
    }
}
