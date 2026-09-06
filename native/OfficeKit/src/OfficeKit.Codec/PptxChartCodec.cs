using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the bounded literal-data p:graphicFrame -> ChartPart projection. The
// chart semantic atoms deliberately reuse the worksheet-chart wire messages;
// PresentationML contributes only its page-relative frame.
internal static partial class PptxChartCodec
{
    private const string ChartGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const int MaxSeries = 256;
    private const int MaxPoints = 1_048_576;
    private static readonly XNamespace ChartNs = ChartGraphicDataUri;
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal sealed record Replacement(
        string PartPath,
        string Sha256,
        bool SlideChanged,
        IReadOnlyCollection<string> ChangedPartPaths,
        IReadOnlyCollection<string> AddedRelationshipKeys,
        IReadOnlyCollection<string> AddedPartPaths,
        IReadOnlyCollection<string> RemovedRelationshipKeys,
        IReadOnlyCollection<string> RemovedPartPaths);

    internal static bool TryRead(P.GraphicFrame source, PptxPartContext context, out PresentationChart chart, out bool editable)
    {
        chart = new PresentationChart();
        editable = false;
        try
        {
            if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is not { Id.HasValue: true } ||
                source.Transform is not { } transform || !TryReadFrame(transform, out var left, out var top, out var width, out var height, out var frameTransform) ||
                source.Graphic?.GraphicData is not { } graphicData ||
                !string.Equals(graphicData.Uri?.Value, ChartGraphicDataUri, StringComparison.Ordinal) ||
                graphicData.Elements<C.ChartReference>().SingleOrDefault()?.Id?.Value is not { Length: > 0 } relationshipId)
                return false;
            ChartPart part;
            try
            {
                if (context.Owner.GetPartById(relationshipId) is not ChartPart chartPart) return false;
                part = chartPart;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
            var xml = ReadXml(part);
            var chartContext = new PptxPartContext(
                part,
                context.SlideIdByPartPath,
                context.SlidePartById,
                context.Assets,
                context.CustomShows);
            if (TryReadComboChart(xml, out chart, out var comboDocument, out editable, allowChartFrameDecorations: true))
            {
                editable &= PptxChartFrameCodec.TryRead(comboDocument.Root!, chartContext, out var comboFrame, out var comboFrameEditable);
                if (comboFrame is not null && (comboFrame.Fill is not null || comboFrame.ImageFill is not null || comboFrame.Line is not null || comboFrame.Shadow is not null))
                {
                    chart.Frame = comboFrame;
                    if (comboFrame.Fill is not null) chart.ChartAreaFill = null;
                }
                editable &= comboFrameEditable;
                editable &= PptxChartTitleTextCodec.TryRead(comboDocument, chart);
                chart.LeftEmu = left;
                chart.TopEmu = top;
                chart.WidthEmu = width;
                chart.HeightEmu = height;
                chart.FrameTransform = frameTransform;
                chart.Accessibility = PptxNonVisualAccessibilityCodec.Read(source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties);
                return true;
            }
            if (!TryReadChart(xml, out var semantic, out var document, out editable)) return false;
            chart = FromSpreadsheet(semantic, left, top, width, height);
            editable &= PptxChartFrameCodec.TryRead(document.Root!, chartContext, out var frame, out var frameEditable);
            if (frame is not null && (frame.Fill is not null || frame.ImageFill is not null || frame.Line is not null || frame.Shadow is not null))
            {
                chart.Frame = frame;
                if (frame.Fill is not null) chart.ChartAreaFill = null;
            }
            editable &= frameEditable;
            editable &= PptxChartTitleTextCodec.TryRead(document, chart);
            chart.FrameTransform = frameTransform;
            chart.Accessibility = PptxNonVisualAccessibilityCodec.Read(source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties);
            return true;
        }
        catch (Exception error) when (error is InvalidOperationException or OverflowException or XmlException)
        {
            chart = new PresentationChart();
            editable = false;
            return false;
        }
    }

    internal static P.GraphicFrame Build(
        PresentationElement element,
        uint nativeId,
        SlidePart slidePart,
        PptxPartContext slideContext)
    {
        Validate(element.Chart, element.Id, element.Name);
        var relationshipId = $"rIdOfficeKitChart{nativeId}";
        var chartPart = slidePart.AddNewPart<ChartPart>(relationshipId);
        var chartContext = new PptxPartContext(
            chartPart,
            slideContext.SlideIdByPartPath,
            slideContext.SlidePartById,
            slideContext.Assets,
            slideContext.CustomShows);
        WriteXml(chartPart, BuildPresentationChartDocument(element.Chart, element.Id, element.Name, chartContext));
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = element.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, element.Chart.Accessibility);
        var transform = new P.Transform(
            new A.Offset { X = element.Chart.LeftEmu, Y = element.Chart.TopEmu },
            new A.Extents { Cx = element.Chart.WidthEmu, Cy = element.Chart.HeightEmu });
        PptxFrameTransformCodec.Apply(transform, element.Chart.FrameTransform);
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                nonVisual,
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            transform,
            new A.Graphic(new A.GraphicData(new C.ChartReference { Id = relationshipId }) { Uri = ChartGraphicDataUri }));
    }

    internal static Replacement Apply(P.GraphicFrame source, PresentationElement requested, PptxPartContext context)
    {
        if (!TryRead(source, context, out var original, out var editable) || !editable)
            throw new CodecException("unsupported_presentation_edit", $"Presentation chart {requested.Id} no longer matches the editable literal-data chart profile.");
        // Source-bound formula references may be edited as references, but
        // source-free authoring remains literal-only. The surrounding
        // topology and formula/cache closure is checked by the shared
        // ChartSpace codec before the part is written.
        Validate(requested.Chart, requested.Id, requested.Name, allowFormulas: true);
        if (!PresentationChartTopologyMatches(requested.Chart, original))
            throw new CodecException("presentation_chart_topology_changed", $"Presentation chart {requested.Id} cannot change chart type, series count, or point topology.");

        var relationshipId = source.Graphic!.GraphicData!.Elements<C.ChartReference>().Single().Id!.Value!;
        var part = (ChartPart)context.Owner.GetPartById(relationshipId);
        var chartContext = new PptxPartContext(
            part,
            context.SlideIdByPartPath,
            context.SlidePartById,
            context.Assets,
            context.CustomShows);
        var document = XDocument.Parse(ReadXml(part), LoadOptions.PreserveWhitespace);
        PatchPresentationChart(document, requested.Chart, requested.Id, requested.Name, chartContext);
        WriteXml(part, document);
        var accessibilityChanged = !object.Equals(requested.Chart.Accessibility, original.Accessibility);
        if (accessibilityChanged)
            PptxNonVisualAccessibilityCodec.ApplyBound(
                source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
                requested.Chart.Accessibility,
                "chart");
        var nonVisual = source.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!;
        var nameChanged = !string.Equals(nonVisual.Name?.Value ?? string.Empty, requested.Name, StringComparison.Ordinal);
        if (nameChanged)
            nonVisual.Name = requested.Name;
        var frameChanged = requested.Chart.LeftEmu != original.LeftEmu || requested.Chart.TopEmu != original.TopEmu ||
            requested.Chart.WidthEmu != original.WidthEmu || requested.Chart.HeightEmu != original.HeightEmu ||
            !object.Equals(requested.Chart.FrameTransform, original.FrameTransform);
        if (frameChanged)
            SetFrame(source.Transform!, requested.Chart);
        var bytes = ReadBytes(part);
        var chartPartPath = Path(part);
        var contextChangedParts = chartContext.RelationshipsChanged
            ? new[] { RelationshipPartPath(part) }
            : Array.Empty<string>();
        var addedPartPaths = chartContext.AddedPartPaths.ToArray();
        var removedPartPaths = chartContext.RemovedPartPaths.ToArray();
        var changedPartPaths = contextChangedParts
            .Concat(addedPartPaths)
            .Concat(removedPartPaths)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return new Replacement(
            chartPartPath,
            Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
            accessibilityChanged || nameChanged || frameChanged,
            changedPartPaths,
            chartContext.AddedRelationshipIds.Select(id => $"{chartPartPath}\0{id}").ToArray(),
            addedPartPaths,
            chartContext.RemovedRelationshipIds.Select(id => $"{chartPartPath}\0{id}").ToArray(),
            removedPartPaths);
    }

    internal static void Validate(PresentationChart? chart, string elementId, string name, bool allowFormulas = false)
    {
        if (chart is null) throw Invalid(elementId, "payload is missing");
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || HasControls(name)) throw Invalid(elementId, "name must contain 1 through 255 characters without controls");
        if (chart.LeftEmu < 0 || chart.TopEmu < 0 || chart.WidthEmu <= 0 || chart.HeightEmu <= 0)
            throw Invalid(elementId, "frame must have non-negative coordinates and positive dimensions");
        PptxFrameTransformCodec.Validate(chart.FrameTransform, elementId, "chart");
        PptxNonVisualAccessibilityCodec.Validate(chart.Accessibility, elementId, "chart");
        PptxChartTitleTextCodec.Validate(chart, elementId);
        if (chart.Type == SpreadsheetChartType.Combo)
        {
            ValidateComboChart(chart, elementId, name, allowFormulas);
            return;
        }
        if (chart.ComboSeries.Count != 0) throw Invalid(elementId, "must not carry combo_series unless type is combo");
        var hasFormulas = chart.Series.Any(series => !string.IsNullOrWhiteSpace(series.CategoryFormula) ||
            !string.IsNullOrWhiteSpace(series.XValueFormula) ||
            !string.IsNullOrWhiteSpace(series.ValueFormula) ||
            !string.IsNullOrWhiteSpace(series.BubbleSizeFormula) ||
            ErrorBarsUseFormula(series));
        if (hasFormulas && !allowFormulas)
            throw Invalid(elementId, "must use literal categories and values without workbook formulas");
        if (hasFormulas && chart.Series.Any(series =>
                !FormulaProfileIsSafe(series.CategoryFormula) ||
                !FormulaProfileIsSafe(series.XValueFormula) ||
                !FormulaProfileIsSafe(series.ValueFormula) ||
                !FormulaProfileIsSafe(series.BubbleSizeFormula) ||
                !FormulaProfileIsSafe(series.ErrorBars?.Plus?.Formula ?? string.Empty) ||
                !FormulaProfileIsSafe(series.ErrorBars?.Minus?.Formula ?? string.Empty)))
            throw Invalid(elementId, "contains a formula outside the local worksheet range profile");
        var spreadsheet = ToSpreadsheet(chart, elementId, name);
        try
        {
            XlsxChartCodec.Validate([spreadsheet], $"presentation/{elementId}", allowStandardRadar: true);
        }
        catch (CodecException error) when (error.Code == "invalid_spreadsheet_chart")
        {
            throw Invalid(elementId, error.Message);
        }
    }

    private static bool ErrorBarsUseFormula(SpreadsheetChartSeriesArtifact series) =>
        !string.IsNullOrWhiteSpace(series.ErrorBars?.Plus?.Formula) ||
        !string.IsNullOrWhiteSpace(series.ErrorBars?.Minus?.Formula);

    internal static void ScrubFrame(P.GraphicFrame source)
    {
        if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            PptxNonVisualAccessibilityCodec.ScrubModeledContent(nonVisual);
            nonVisual.Name = string.Empty;
        }
        if (source.Transform is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
            PptxFrameTransformCodec.Scrub(transform);
        }
    }

    private static SpreadsheetChartArtifact ToSpreadsheet(PresentationChart source, string id, string name)
    {
        var output = new SpreadsheetChartArtifact
        {
            Id = id,
            Name = name,
            Title = source.Title,
            Type = source.Type,
            HasLegend = source.HasLegend,
            LegendPosition = source.LegendPosition,
            Grouping = source.Grouping,
            BarDirection = source.BarDirection,
            AbsoluteAnchor = new SpreadsheetAbsoluteAnchorArtifact
            {
                XEmu = source.LeftEmu,
                YEmu = source.TopEmu,
                WidthEmu = source.WidthEmu,
                HeightEmu = source.HeightEmu,
            },
        };
        output.Categories.Add(source.Categories);
        output.Series.Add(source.Series.Select(series => series.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.HasShowCategoryAxis)
        {
            output.XAxis ??= new SpreadsheetChartAxisArtifact();
            output.XAxis.Visible = source.ShowCategoryAxis;
        }
        if (source.HasShowValueAxis)
        {
            output.YAxis ??= new SpreadsheetChartAxisArtifact();
            output.YAxis.Visible = source.ShowValueAxis;
        }
        if (source.HasShowGridlines)
        {
            output.YAxis ??= new SpreadsheetChartAxisArtifact();
            output.YAxis.ShowMajorGridlines = source.ShowGridlines;
        }
        if (source.HasGapWidth) output.GapWidth = source.GapWidth;
        if (source.HasOverlap) output.Overlap = source.Overlap;
        if (source.HasVaryColors) output.VaryColors = source.VaryColors;
        if (source.ChartAreaFill is not null) output.ChartAreaFill = source.ChartAreaFill.Clone();
        if (source.PlotAreaFill is not null) output.PlotAreaFill = source.PlotAreaFill.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        if (source.TitleTextStyle is not null) output.TitleTextStyle = source.TitleTextStyle.Clone();
        if (source.LegendTextStyle is not null) output.LegendTextStyle = source.LegendTextStyle.Clone();
        if (source.LegendFill is not null) output.LegendFill = source.LegendFill.Clone();
        if (source.LegendLine is not null) output.LegendLine = source.LegendLine.Clone();
        if (source.LineOptions is not null) output.LineOptions = source.LineOptions.Clone();
        if (source.HasFirstSliceAngle) output.FirstSliceAngle = source.FirstSliceAngle;
        if (source.HasDoughnutHoleSize) output.DoughnutHoleSize = source.DoughnutHoleSize;
        if (source.HasBubbleScale) output.BubbleScale = source.BubbleScale;
        output.BubbleSizeMode = source.BubbleSizeMode;
        if (source.HasDisplayBlanksAs) output.DisplayBlanksAs = source.DisplayBlanksAs;
        if (source.HasLegendOverlay) output.LegendOverlay = source.LegendOverlay;
        return output;
    }

    private static PresentationChart FromSpreadsheet(SpreadsheetChartArtifact source, long left, long top, long width, long height)
    {
        var output = new PresentationChart
        {
            LeftEmu = left,
            TopEmu = top,
            WidthEmu = width,
            HeightEmu = height,
            Type = source.Type,
            Title = source.Title,
            HasLegend = source.HasLegend,
            LegendPosition = source.LegendPosition,
            Grouping = source.Grouping,
            BarDirection = source.BarDirection,
        };
        output.Categories.Add(source.Categories);
        output.Series.Add(source.Series.Select(series => series.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.XAxis?.HasVisible == true) output.ShowCategoryAxis = source.XAxis.Visible;
        if (source.YAxis?.HasVisible == true) output.ShowValueAxis = source.YAxis.Visible;
        if (source.YAxis?.HasShowMajorGridlines == true) output.ShowGridlines = source.YAxis.ShowMajorGridlines;
        if (source.HasGapWidth) output.GapWidth = source.GapWidth;
        if (source.HasOverlap) output.Overlap = source.Overlap;
        if (source.HasVaryColors) output.VaryColors = source.VaryColors;
        if (source.ChartAreaFill is not null) output.ChartAreaFill = source.ChartAreaFill.Clone();
        if (source.PlotAreaFill is not null) output.PlotAreaFill = source.PlotAreaFill.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        if (source.TitleTextStyle is not null) output.TitleTextStyle = source.TitleTextStyle.Clone();
        if (source.LegendTextStyle is not null) output.LegendTextStyle = source.LegendTextStyle.Clone();
        if (source.LegendFill is not null) output.LegendFill = source.LegendFill.Clone();
        if (source.LegendLine is not null) output.LegendLine = source.LegendLine.Clone();
        if (source.LineOptions is not null) output.LineOptions = source.LineOptions.Clone();
        if (source.HasFirstSliceAngle) output.FirstSliceAngle = source.FirstSliceAngle;
        if (source.HasDoughnutHoleSize) output.DoughnutHoleSize = source.DoughnutHoleSize;
        if (source.HasBubbleScale) output.BubbleScale = source.BubbleScale;
        output.BubbleSizeMode = source.BubbleSizeMode;
        if (source.HasDisplayBlanksAs) output.DisplayBlanksAs = source.DisplayBlanksAs;
        if (source.HasLegendOverlay) output.LegendOverlay = source.LegendOverlay;
        return output;
    }

    private static bool TryReadChart(string xml, out SpreadsheetChartArtifact chart, out XDocument document, out bool editable)
    {
        if (!OpenXmlChartSpaceCodec.TryRead(xml, out chart, out document, out editable, allowRichTitle: true, allowChartFrameDecorations: true)) return false;
        return chart.Series.All(series => FormulaProfileIsSafe(series.CategoryFormula) &&
            FormulaProfileIsSafe(series.XValueFormula) &&
            FormulaProfileIsSafe(series.ValueFormula) &&
            FormulaProfileIsSafe(series.BubbleSizeFormula));
    }

    internal static bool FormulaProfileIsSafe(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula)) return true;
        // Keep the source-bound profile on one local worksheet range. An
        // external [book.xlsx] link, error token, structured reference,
        // defined name, function, union, or 3-D reference cannot be edited
        // without owning the workbook relationship/evaluation graph.
        return LocalWorksheetRangeFormula().IsMatch(formula);
    }

    [GeneratedRegex("^(?:[A-Za-z_][A-Za-z0-9_.]*|'(?:[^']|'')+')!\\$?[A-Z]{1,3}\\$?[1-9][0-9]*(?::\\$?[A-Z]{1,3}\\$?[1-9][0-9]*)?$", RegexOptions.CultureInvariant)]
    private static partial Regex LocalWorksheetRangeFormula();

    private static XDocument BuildChartDocument(SpreadsheetChartArtifact chart)
    {
        return OpenXmlChartSpaceCodec.Build(chart);
    }

    private static void PatchChart(XDocument document, SpreadsheetChartArtifact target, bool patchTitle)
    {
        OpenXmlChartSpaceCodec.Patch(
            document,
            target,
            "presentation_chart_topology_changed",
            "Presentation chart",
            patchTitle,
            allowChartFrameDecorations: true);
    }

    private static bool TryReadFrame(
        P.Transform transform,
        out long left,
        out long top,
        out long width,
        out long height,
        out PresentationFrameTransform? frameTransform)
    {
        left = top = width = height = 0;
        frameTransform = null;
        if (transform.ChildElements.Count != 2 ||
            transform.Offset?.X?.Value is null || transform.Offset.Y?.Value is null ||
            transform.Extents?.Cx?.Value is null or <= 0 || transform.Extents.Cy?.Value is null or <= 0 ||
            transform.Offset.X.Value < 0 || transform.Offset.Y.Value < 0 ||
            !PptxFrameTransformCodec.TryRead(transform, out frameTransform)) return false;
        left = transform.Offset.X.Value; top = transform.Offset.Y.Value; width = transform.Extents.Cx.Value; height = transform.Extents.Cy.Value;
        return true;
    }

    private static void SetFrame(P.Transform transform, PresentationChart chart)
    {
        transform.Offset!.X = chart.LeftEmu; transform.Offset.Y = chart.TopEmu;
        transform.Extents!.Cx = chart.WidthEmu; transform.Extents.Cy = chart.HeightEmu;
        PptxFrameTransformCodec.Apply(transform, chart.FrameTransform);
    }

    private static string ReadXml(OpenXmlPart part) => Encoding.UTF8.GetString(ReadBytes(part));
    private static byte[] ReadBytes(OpenXmlPart part) { using var stream = part.GetStream(FileMode.Open, FileAccess.Read); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
    private static void WriteXml(OpenXmlPart part, XDocument document) { using var stream = part.GetStream(FileMode.Create, FileAccess.Write); using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false, Indent = false }); document.Save(writer); }
    private static string Path(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string RelationshipPartPath(OpenXmlPart part)
    {
        var path = Path(part);
        var separator = path.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : path[..separator];
        var fileName = separator < 0 ? path : path[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException Invalid(string id, string message) => new("invalid_presentation_chart", $"Presentation chart {id} {message}.");
}
