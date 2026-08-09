using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed class OpenXmlChartErrorBarsCodecTests
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    [Fact]
    public void SharedChartSpaceErrorBarsBuildReadPatchAndFailClosed()
    {
        var chart = Chart();
        chart.Series[1].ErrorBars.Direction = SpreadsheetChartErrorBarDirection.X;
        XlsxChartCodec.Validate([chart], "summary");

        var document = OpenXmlChartSpaceCodec.Build(chart);
        var native = document.Descendants(ChartNs + "errBars").ToArray();
        Assert.Equal(5, native.Length);
        Assert.Equal(["cust", "fixedVal", "percentage", "stdDev", "stdErr"], native.Select(item => (string)item.Element(ChartNs + "errValType")!.Attribute("val")!));
        Assert.Equal(["y", "x", "y", "y", "y"], native.Select(item => (string)item.Element(ChartNs + "errDir")!.Attribute("val")!));
        Assert.Equal(["errDir", "errBarType", "errValType", "noEndCap", "plus", "minus", "spPr"], native[0].Elements().Select(item => item.Name.LocalName));
        Assert.Equal("'Summary'!$D$2:$D$4", native[0].Element(ChartNs + "minus")!.Descendants(ChartNs + "f").Single().Value);
        Assert.Equal("2", native[1].Element(ChartNs + "val")!.Attribute("val")!.Value);
        Assert.Null(native[4].Element(ChartNs + "val"));

        Assert.True(OpenXmlChartSpaceCodec.TryRead(document.ToString(SaveOptions.DisableFormatting), out var imported, out var editableDocument, out var editable));
        Assert.True(editable);
        Assert.Equal(5, imported.Series.Count);
        var custom = imported.Series[0].ErrorBars;
        Assert.NotNull(custom);
        Assert.Equal(SpreadsheetChartErrorBarValueType.Custom, custom.ValueType);
        Assert.Equal([1D, 2D, 3D], custom.Plus.Values);
        Assert.Equal("'Summary'!$D$2:$D$4", custom.Minus.Formula);
        Assert.Equal("0.0", custom.Minus.FormatCode);
        Assert.Equal("7C3AED", custom.Line.Color.Rgb);

        var fixedValue = imported.Series[1].ErrorBars;
        Assert.Equal(SpreadsheetChartErrorBarDirection.X, fixedValue.Direction);
        fixedValue.Value = 3.5;
        fixedValue.NoEndCap = true;
        fixedValue.Line = new SpreadsheetChartLineStyleArtifact { Color = new SpreadsheetColor { Rgb = "0EA5E9" }, WidthPoints = 2 };
        OpenXmlChartSpaceCodec.Patch(editableDocument, imported, "error_bar_topology_changed", "Test chart");
        Assert.True(OpenXmlChartSpaceCodec.TryRead(editableDocument.ToString(SaveOptions.DisableFormatting), out var patched, out _, out var patchedEditable));
        Assert.True(patchedEditable);
        Assert.Equal(3.5, patched.Series[1].ErrorBars.Value);
        Assert.True(patched.Series[1].ErrorBars.NoEndCap);
        Assert.Equal("0EA5E9", patched.Series[1].ErrorBars.Line.Color.Rgb);

        var unsupported = new XDocument(document);
        unsupported.Descendants(ChartNs + "errBars").First().Add(new XElement(ChartNs + "extLst"));
        Assert.True(OpenXmlChartSpaceCodec.TryRead(unsupported.ToString(SaveOptions.DisableFormatting), out var preserved, out _, out var unsupportedEditable));
        Assert.False(unsupportedEditable);
        Assert.Null(preserved.Series[0].ErrorBars);
        Assert.NotNull(preserved.Series[1].ErrorBars);

        var topologyChanged = imported.Clone();
        topologyChanged.Series[0].ErrorBars = null;
        var error = Assert.Throws<CodecException>(() => OpenXmlChartSpaceCodec.Patch(editableDocument, topologyChanged, "error_bar_topology_changed", "Test chart"));
        Assert.Equal("error_bar_topology_changed", error.Code);
    }

    [Fact]
    public void ErrorBarValidationRejectsUnsupportedAndContradictoryState()
    {
        var area = Chart();
        area.Type = SpreadsheetChartType.Area;
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([area], "summary")).Code);

        var missingSide = Chart();
        missingSide.Series[0].ErrorBars.Plus = null;
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([missingSide], "summary")).Code);

        var scalarStandardError = Chart();
        scalarStandardError.Series[4].ErrorBars.Value = 1;
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([scalarStandardError], "summary")).Code);

        var nonCustomData = Chart();
        nonCustomData.Series[1].ErrorBars.Plus = new SpreadsheetChartErrorBarDataArtifact { Values = { 1, 1, 1 } };
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([nonCustomData], "summary")).Code);

        var invalidFormula = Chart();
        invalidFormula.Series[0].ErrorBars.Minus.Formula = "='Summary'!$D$2:$D$4";
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([invalidFormula], "summary")).Code);
    }

    private static SpreadsheetChartArtifact Chart()
    {
        var chart = new SpreadsheetChartArtifact
        {
            Id = "chart/error-bars",
            Name = "Error-bar chart",
            Title = "Uncertainty evidence",
            Type = SpreadsheetChartType.Line,
            HasLegend = true,
            AbsoluteAnchor = new SpreadsheetAbsoluteAnchorArtifact { XEmu = 0, YEmu = 0, WidthEmu = 4_000_000, HeightEmu = 2_500_000 },
            XAxis = new SpreadsheetChartAxisArtifact { Title = "Period" },
            YAxis = new SpreadsheetChartAxisArtifact { Title = "Value" },
        };
        chart.Categories.Add(["A", "B", "C"]);
        chart.Series.Add(Series("Custom", new SpreadsheetChartErrorBarsArtifact
        {
            Direction = SpreadsheetChartErrorBarDirection.Y,
            Type = SpreadsheetChartErrorBarType.Both,
            ValueType = SpreadsheetChartErrorBarValueType.Custom,
            NoEndCap = true,
            Plus = new SpreadsheetChartErrorBarDataArtifact { Values = { 1, 2, 3 }, FormatCode = "0.0" },
            Minus = new SpreadsheetChartErrorBarDataArtifact { Formula = "'Summary'!$D$2:$D$4", Values = { 0.5, 1, 1.5 }, FormatCode = "0.0" },
            Line = new SpreadsheetChartLineStyleArtifact { Color = new SpreadsheetColor { Rgb = "7C3AED" }, DashStyle = SpreadsheetChartLineDashStyle.Dashed, WidthPoints = 1.5 },
        }));
        chart.Series.Add(Series("Fixed", Scalar(SpreadsheetChartErrorBarValueType.FixedValue, 2)));
        chart.Series.Add(Series("Percentage", Scalar(SpreadsheetChartErrorBarValueType.Percentage, 5)));
        chart.Series.Add(Series("Deviation", Scalar(SpreadsheetChartErrorBarValueType.StandardDeviation, 1.5)));
        chart.Series.Add(Series("Standard error", new SpreadsheetChartErrorBarsArtifact
        {
            Direction = SpreadsheetChartErrorBarDirection.Y,
            Type = SpreadsheetChartErrorBarType.Both,
            ValueType = SpreadsheetChartErrorBarValueType.StandardError,
        }));
        return chart;
    }

    private static SpreadsheetChartSeriesArtifact Series(string name, SpreadsheetChartErrorBarsArtifact errorBars)
    {
        var series = new SpreadsheetChartSeriesArtifact { Name = name, ErrorBars = errorBars };
        series.Values.Add([4, 7, 11]);
        return series;
    }

    private static SpreadsheetChartErrorBarsArtifact Scalar(SpreadsheetChartErrorBarValueType valueType, double value) => new()
    {
        Direction = SpreadsheetChartErrorBarDirection.Y,
        Type = SpreadsheetChartErrorBarType.Both,
        ValueType = valueType,
        Value = value,
    };
}
