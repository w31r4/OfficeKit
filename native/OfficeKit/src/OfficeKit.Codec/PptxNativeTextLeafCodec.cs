using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxNativeTextLeaf(uint Index, string Text, A.Text Element);

// Exposes only the text tokens of an otherwise opaque DrawingML table.  The
// table keeps its original XML, styles, runs, and topology; an issued leaf
// can replace one existing a:t token without asking the semantic table model
// to guess how rich text should be represented.
internal static class PptxNativeTextLeafCodec
{
    private const string TableGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/table";
    private const int MaxLeaves = 4_096;
    private const int MaxLeafLength = 32_767;

    internal static bool TryDescribe(OpenXmlElement source, out IReadOnlyList<PptxNativeTextLeaf> leaves)
    {
        leaves = Array.Empty<PptxNativeTextLeaf>();
        var texts = source switch
        {
            P.GraphicFrame frame when frame.Graphic?.GraphicData?.Uri?.Value == TableGraphicDataUri =>
                DescribeTable(frame),
            P.GroupShape group => DescribeGroup(group),
            _ => Array.Empty<A.Text>(),
        };
        if (texts.Length is < 1 or > MaxLeaves) return false;

        leaves = texts.Select((text, index) => new PptxNativeTextLeaf(checked((uint)index), text.Text, text)).ToArray();
        return true;
    }

    private static A.Text[] DescribeTable(P.GraphicFrame frame)
    {
        var tables = frame.Descendants<A.Table>().ToArray();
        if (tables.Length != 1) return Array.Empty<A.Text>();
        var cells = tables[0].Descendants<A.TableCell>().ToArray();
        var texts = tables[0].Descendants<A.Text>().ToArray();
        if (cells.Length == 0) return Array.Empty<A.Text>();

        // Every exposed token must be a direct run text inside one table cell.
        // This excludes fields and foreign text-bearing extensions whose
        // formatting/meaning would not be source-bound by this profile.
        return texts.All(text => text.Parent is A.Run && text.Ancestors<A.TableCell>().Count() == 1 && ValidText(text.Text))
            ? texts
            : Array.Empty<A.Text>();
    }

    private static A.Text[] DescribeGroup(P.GroupShape group)
    {
        // Empty a:t placeholders are not emitted by the bounded JS leaf index;
        // they carry no visible value and remain part of the opaque source.
        var texts = group.Descendants<A.Text>().Where(text => !string.IsNullOrEmpty(text.Text)).ToArray();
        // A group can contain opaque geometry, images, and nested groups.  A
        // token is safe to expose only when it belongs to exactly one regular
        // shape/run; table, chart, field, and extension text stays opaque.
        return texts.All(text => text.Parent is A.Run && text.Ancestors<P.Shape>().Count() == 1 &&
                                 text.Ancestors<A.TableCell>().Count() == 0 && ValidText(text.Text))
            ? texts
            : Array.Empty<A.Text>();
    }

    internal static bool TryResolve(OpenXmlElement source, uint index, out PptxNativeTextLeaf leaf)
    {
        leaf = null!;
        if (!TryDescribe(source, out var leaves) || index >= (uint)leaves.Count) return false;
        leaf = leaves[(int)index];
        return true;
    }

    private static bool ValidText(string value)
    {
        if (value.Length > MaxLeafLength) return false;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (character <= '\b' || character is '\v' or '\f' || character is >= '\u000E' and <= '\u001F')
                return false;
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
