using System.Xml.Linq;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns the bounded Presentation chart c:title/c:tx/c:rich profile. The
// paragraph/run semantics reuse PresentationTextBody; formula titles and
// external title hyperlinks remain source-owned.
internal static class PptxChartTitleTextCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static bool TryRead(XDocument document, PresentationChart chart)
    {
        var title = document.Root?.Element(ChartNs + "chart")?.Element(ChartNs + "title");
        if (title is null) return chart.Title.Length == 0;
        var children = title.Elements().ToArray();
        if (children.Any(child => child.Name != ChartNs + "tx" && child.Name != ChartNs + "layout" && child.Name != ChartNs + "overlay") ||
            children.Count(child => child.Name == ChartNs + "tx") != 1) return false;
        var tx = title.Element(ChartNs + "tx")!;
        var rich = tx.Element(ChartNs + "rich");
        if (rich is null || tx.Elements().Count() != 1 ||
            rich.Descendants(DrawingNs + "hlinkClick").Any() ||
            rich.Descendants(DrawingNs + "hlinkMouseOver").Any()) return false;
        var richChildren = rich.Elements().ToArray();
        if (richChildren.Length < 3 || richChildren[0].Name != DrawingNs + "bodyPr" ||
            richChildren[1].Name != DrawingNs + "lstStyle" ||
            richChildren.Skip(2).Any(child => child.Name != DrawingNs + "p")) return false;

        A.TextBody native;
        try
        {
            native = new A.TextBody { InnerXml = string.Concat(rich.Nodes().Select(NodeXml)) };
        }
        catch (Exception error) when (error is InvalidOperationException or System.Xml.XmlException)
        {
            return false;
        }
        if (!PptxTextCodec.SupportsEditing(native)) return false;
        var body = PptxTextCodec.ReadDrawingTextBody(native);
        if (!PptxTextCodec.Flatten(body).Equals(chart.Title, StringComparison.Ordinal)) return false;

        var styleProbe = new SpreadsheetChartArtifact();
        if (!XlsxChartTextStyleCodec.TryReadTitle(title, styleProbe)) chart.TitleBody = body;
        return true;
    }

    internal static void Validate(PresentationChart chart, string elementId)
    {
        if (chart.TitleBody is null) return;
        if (!PptxTextCodec.Flatten(chart.TitleBody).Equals(chart.Title, StringComparison.Ordinal))
            throw Invalid(elementId, "structured title body must flatten to title");
        if (chart.TitleBody.Paragraphs.Count == 0 || chart.Title.Length == 0)
            throw Invalid(elementId, "structured title must contain visible text");
        if (chart.TitleBody.Paragraphs.SelectMany(paragraph => paragraph.Runs).Any(run =>
                run.HyperlinkCase != PresentationTextRun.HyperlinkOneofCase.None))
            throw Invalid(elementId, "structured title hyperlinks are not part of the bounded ChartPart profile");
        PptxTextCodec.Validate(new PresentationShape
        {
            Text = chart.Title,
            TextBody = chart.TitleBody.Clone(),
        });
    }

    internal static void Apply(XDocument document, PresentationChart chart)
    {
        if (chart.TitleBody is null) return;
        var nativeChart = document.Root?.Element(ChartNs + "chart") ??
            throw new CodecException("unsupported_presentation_edit", "Presentation chart is missing c:chart.");
        var title = nativeChart.Element(ChartNs + "title");
        if (title is null)
        {
            title = new XElement(ChartNs + "title", new XElement(ChartNs + "layout"));
            var plotArea = nativeChart.Element(ChartNs + "plotArea") ??
                throw new CodecException("unsupported_presentation_edit", "Presentation chart is missing c:plotArea.");
            plotArea.AddBeforeSelf(title);
        }
        var existingTx = title.Element(ChartNs + "tx");
        var replacement = new XElement(ChartNs + "tx", RichText(chart.TitleBody));
        if (existingTx is null) title.AddFirst(replacement);
        else existingTx.ReplaceWith(replacement);
        OpenXmlChartSpaceCodec.PatchTitlePlacement(
            title,
            chart.HasTitlePlacement ? chart.TitlePlacement : string.Empty,
            "unsupported_presentation_edit",
            "Presentation chart title");
    }

    internal static bool RequiresPlainRewrite(XDocument document)
    {
        var title = document.Root?.Element(ChartNs + "chart")?.Element(ChartNs + "title");
        if (title is null) return false;
        return !XlsxChartTextStyleCodec.TryReadTitle(title, new SpreadsheetChartArtifact());
    }

    internal static void ApplyPlain(XDocument document, PresentationChart chart)
    {
        var nativeChart = document.Root?.Element(ChartNs + "chart") ??
            throw new CodecException("unsupported_presentation_edit", "Presentation chart is missing c:chart.");
        var title = nativeChart.Element(ChartNs + "title");
        if (chart.Title.Length == 0)
        {
            title?.Remove();
            return;
        }
        var replacement = XlsxChartTextStyleCodec.TitleElement(
            chart.Title,
            chart.TitleTextStyle,
            chart.HasTitlePlacement ? chart.TitlePlacement : string.Empty);
        if (title is null)
        {
            var plotArea = nativeChart.Element(ChartNs + "plotArea") ??
                throw new CodecException("unsupported_presentation_edit", "Presentation chart is missing c:plotArea.");
            plotArea.AddBeforeSelf(replacement);
            return;
        }
        var existingTx = title.Element(ChartNs + "tx") ??
            throw new CodecException("unsupported_presentation_edit", "Presentation chart title is missing c:tx.");
        existingTx.ReplaceWith(replacement.Element(ChartNs + "tx")!);
        OpenXmlChartSpaceCodec.PatchTitlePlacement(
            title,
            chart.HasTitlePlacement ? chart.TitlePlacement : string.Empty,
            "unsupported_presentation_edit",
            "Presentation chart title");
    }

    private static XElement RichText(PresentationTextBody body)
    {
        var native = PptxTextCodec.BuildDrawingTextBody(body);
        return new XElement(ChartNs + "rich",
            native.ChildElements.Select(child => XElement.Parse(child.OuterXml, LoadOptions.PreserveWhitespace)));
    }

    private static string NodeXml(XNode node) => node.ToString(SaveOptions.DisableFormatting);

    private static CodecException Invalid(string id, string message) =>
        new("invalid_presentation_chart", $"Presentation chart {id} {message}.");
}
