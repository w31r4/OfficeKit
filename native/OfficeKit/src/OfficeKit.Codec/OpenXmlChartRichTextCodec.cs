using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static class OpenXmlChartRichTextCodec
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartRichTextArtifact? text)
    {
        if (text is null) return;
        if (text.Paragraphs.Count is 0 or > 4096 || text.Paragraphs.Sum(p => (long)p.Runs.Count) > 16384 ||
            text.Paragraphs.SelectMany(p => p.Runs).Sum(r => (long)r.Text.Length) > 1_048_576)
            throw Invalid("Rich chart text exceeds paragraph/inline/text budgets or has no paragraphs.");
        foreach (var paragraph in text.Paragraphs)
        {
            Style(paragraph.Style, character: false);
            Style(paragraph.EndStyle, character: true);
            foreach (var run in paragraph.Runs)
            {
                if (run.ContentCase == SpreadsheetChartTextRunArtifact.ContentOneofCase.Text)
                {
                    if (run.Text.Any(char.IsControl)) throw Invalid("Use typed chart text breaks instead of control characters.");
                }
                else if (run.ContentCase != SpreadsheetChartTextRunArtifact.ContentOneofCase.LineBreak || !run.LineBreak)
                    throw Invalid("A chart text run requires text or a true line break.");
                Style(run.Style, character: true);
            }
        }
    }

    private static void Style(SpreadsheetChartTextStyleArtifact? style, bool character)
    {
        XlsxChartTextStyleCodec.ValidateStyle(style, "chart", "chart", "rich text style");
        if (character && style is { Alignment.Length: > 0 }) throw Invalid("Chart character styles cannot carry paragraph alignment.");
    }

    internal static void NormalizeLabel(SpreadsheetChartTrendlineLabelArtifact label)
    {
        if (label.RichText is not { Paragraphs.Count: 1 } rich) return;
        var paragraph = rich.Paragraphs[0];
        if (paragraph.Style is not null || paragraph.EndStyle is not null || paragraph.Runs.Count != 1) return;
        var run = paragraph.Runs[0];
        if (run.Style is not null || run.ContentCase != SpreadsheetChartTextRunArtifact.ContentOneofCase.Text || run.Text.Length is 0 or > 255) return;
        label.Text = run.Text;
        label.RichText = null;
    }

    internal static XElement Element(SpreadsheetChartRichTextArtifact text)
    {
        Validate(text);
        return new XElement(C + "tx", new XElement(C + "rich", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"),
            text.Paragraphs.Select(p => new XElement(A + "p",
                p.Style is null ? null : XlsxChartTextStyleCodec.ParagraphProperties(p.Style),
                p.Runs.Select(r => new XElement(A + (r.ContentCase == SpreadsheetChartTextRunArtifact.ContentOneofCase.LineBreak ? "br" : "r"),
                    r.Style is null ? null : XlsxChartTextStyleCodec.StyleProperties("rPr", r.Style),
                    r.ContentCase == SpreadsheetChartTextRunArtifact.ContentOneofCase.Text ? new XElement(A + "t", r.Text) : null)),
                p.EndStyle is null ? null : XlsxChartTextStyleCodec.StyleProperties("endParaRPr", p.EndStyle)))));
    }

    internal static bool TryRead(XElement source, out SpreadsheetChartRichTextArtifact text)
    {
        text = new SpreadsheetChartRichTextArtifact();
        if (source.Name != C + "tx" || !NoAttributes(source) || source.Elements().Count() != 1 ||
            source.DescendantNodes().Any(node => node is not (XElement or XText) ||
                node is XText literal && literal.Parent?.Name != A + "t" && !string.IsNullOrWhiteSpace(literal.Value))) return false;
        var rich = source.Elements().Single();
        if (rich.Name != C + "rich" || !NoAttributes(rich)) return false;
        var children = rich.Elements().ToArray();
        if (children.Length is < 3 or > 4098 || children[0].Name != A + "bodyPr" || !Empty(children[0]) ||
            children[1].Name != A + "lstStyle" || !Empty(children[1])) return false;
        foreach (var p in children.Skip(2))
        {
            if (p.Name != A + "p" || !NoAttributes(p)) return false;
            var paragraph = new SpreadsheetChartTextParagraphArtifact();
            var items = p.Elements().ToArray();
            var i = 0;
            if (i < items.Length && items[i].Name == A + "pPr")
            {
                var properties = new XElement(items[i++]);
                if (properties.Elements(A + "defRPr").Count() > 1) return false;
                properties.Elements(A + "defRPr").Where(Empty).Remove();
                if (!Empty(properties))
                {
                    if (!XlsxChartTextStyleCodec.TryReadParagraphStyle(properties, out var style)) return false;
                    paragraph.Style = style;
                }
            }
            while (i < items.Length && (items[i].Name == A + "r" || items[i].Name == A + "br"))
            {
                var inline = items[i++];
                if (!NoAttributes(inline)) return false;
                var run = new SpreadsheetChartTextRunArtifact();
                var nodes = inline.Elements().ToArray();
                var j = 0;
                if (j < nodes.Length && nodes[j].Name == A + "rPr")
                {
                    if (!ReadCharacterStyle(nodes[j++], out var style)) return false;
                    run.Style = style;
                }
                if (inline.Name == A + "br") run.LineBreak = true;
                else
                {
                    if (j >= nodes.Length || nodes[j].Name != A + "t" || nodes[j].HasElements ||
                        nodes[j].Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name != XNamespace.Xml + "space" || a.Value is not ("preserve" or "default")))) return false;
                    run.Text = nodes[j++].Value;
                }
                if (j != nodes.Length) return false;
                paragraph.Runs.Add(run);
            }
            if (i < items.Length && items[i].Name == A + "endParaRPr")
            {
                if (!ReadCharacterStyle(items[i++], out var style)) return false;
                paragraph.EndStyle = style;
            }
            if (i != items.Length) return false;
            text.Paragraphs.Add(paragraph);
        }
        try { Validate(text); return true; }
        catch (CodecException) { return false; }
    }

    private static bool ReadCharacterStyle(XElement source, out SpreadsheetChartTextStyleArtifact? style)
    {
        style = null;
        if (Empty(source)) return true;
        if (!XlsxChartTextStyleCodec.TryExactStyleProperties(source, out var parsed)) return false;
        style = parsed;
        return true;
    }

    internal static SpreadsheetChartRichTextArtifact FromPpj(JsonElement source,
        Func<JsonElement, string> resolveText, Func<JsonElement, SpreadsheetChartTextStyleArtifact> resolveStyle)
    {
        var text = new SpreadsheetChartRichTextArtifact();
        foreach (var p in source.GetProperty("paragraphs").EnumerateArray())
        {
            var paragraph = new SpreadsheetChartTextParagraphArtifact();
            if (p.TryGetProperty("style", out var style)) paragraph.Style = resolveStyle(style);
            if (p.TryGetProperty("endStyle", out var end)) paragraph.EndStyle = resolveStyle(end);
            foreach (var r in p.GetProperty("runs").EnumerateArray())
            {
                var run = new SpreadsheetChartTextRunArtifact();
                if (r.TryGetProperty("text", out var literal)) run.Text = resolveText(literal);
                if (r.TryGetProperty("break", out var br))
                {
                    if (run.ContentCase != SpreadsheetChartTextRunArtifact.ContentOneofCase.None) throw Invalid("Chart text cannot combine text and break.");
                    run.LineBreak = br.GetBoolean();
                }
                if (r.TryGetProperty("style", out var runStyle)) run.Style = resolveStyle(runStyle);
                paragraph.Runs.Add(run);
            }
            text.Paragraphs.Add(paragraph);
        }
        Validate(text);
        return text;
    }

    internal static JsonObject Project(SpreadsheetChartRichTextArtifact text, Func<SpreadsheetChartTextStyleArtifact, JsonObject> projectStyle)
    {
        var paragraphs = new JsonArray();
        foreach (var p in text.Paragraphs)
        {
            var paragraph = new JsonObject();
            if (p.Style is not null) paragraph["style"] = projectStyle(p.Style);
            var runs = new JsonArray();
            foreach (var r in p.Runs)
            {
                var run = new JsonObject();
                if (r.ContentCase == SpreadsheetChartTextRunArtifact.ContentOneofCase.Text) run["text"] = r.Text;
                else run["break"] = true;
                if (r.Style is not null) run["style"] = projectStyle(r.Style);
                runs.Add(run);
            }
            paragraph["runs"] = runs;
            if (p.EndStyle is not null) paragraph["endStyle"] = projectStyle(p.EndStyle);
            paragraphs.Add(paragraph);
        }
        return new JsonObject { ["paragraphs"] = paragraphs };
    }

    private static bool NoAttributes(XElement element) => element.Attributes().All(a => a.IsNamespaceDeclaration);
    private static bool Empty(XElement element) => NoAttributes(element) && !element.HasElements;
    private static CodecException Invalid(string message) => new("invalid_spreadsheet_chart", message);
}
