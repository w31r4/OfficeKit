using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed class OpenXmlChartTrendlineCodecTests
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    [Fact]
    public void SharedChartSpaceTrendlinesBuildReadPatchAndFailClosed()
    {
        var chart = Chart();
        XlsxChartCodec.Validate([chart], "summary");

        var document = OpenXmlChartSpaceCodec.Build(chart);
        var native = document.Descendants(ChartNs + "trendline").ToArray();
        Assert.Equal(6, native.Length);
        Assert.Equal(["exp", "linear", "log", "movingAvg", "poly", "power"], native.Select(item => (string)item.Element(ChartNs + "trendlineType")!.Attribute("val")!));
        Assert.Equal("3", native[3].Element(ChartNs + "period")!.Attribute("val")!.Value);
        Assert.Equal("3", native[4].Element(ChartNs + "order")!.Attribute("val")!.Value);
        Assert.Equal("0.5", native[1].Element(ChartNs + "forward")!.Attribute("val")!.Value);
        Assert.Equal("1", native[1].Element(ChartNs + "backward")!.Attribute("val")!.Value);
        Assert.Equal("0", native[1].Element(ChartNs + "intercept")!.Attribute("val")!.Value);
        Assert.NotNull(native[1].Element(ChartNs + "dispEq"));
        Assert.NotNull(native[1].Element(ChartNs + "dispRSqr"));

        Assert.True(OpenXmlChartSpaceCodec.TryRead(document.ToString(SaveOptions.DisableFormatting), out var imported, out var editableDocument, out var editable));
        Assert.True(editable);
        var trendlines = Assert.Single(imported.Series).Trendlines;
        Assert.Equal(6, trendlines.Count);
        Assert.Equal(SpreadsheetChartTrendlineType.MovingAverage, trendlines[3].Type);
        Assert.True(trendlines[4].HasPolynomialOrder);
        Assert.Equal(3U, trendlines[4].PolynomialOrder);
        Assert.Equal("7C3AED", trendlines[1].Line.Color.Rgb);

        trendlines[1].Name = "Edited linear evidence";
        trendlines[1].Forward = 1.5;
        trendlines[1].Line.Color.Rgb = "0EA5E9";
        OpenXmlChartSpaceCodec.Patch(editableDocument, imported, "trendline_topology_changed", "Test chart");
        Assert.True(OpenXmlChartSpaceCodec.TryRead(editableDocument.ToString(SaveOptions.DisableFormatting), out var patched, out _, out var patchedEditable));
        Assert.True(patchedEditable);
        Assert.Equal("Edited linear evidence", patched.Series[0].Trendlines[1].Name);
        Assert.Equal(1.5, patched.Series[0].Trendlines[1].Forward);
        Assert.Equal("0EA5E9", patched.Series[0].Trendlines[1].Line.Color.Rgb);

        var unsupported = new XDocument(document);
        unsupported.Descendants(ChartNs + "trendline").First().Add(new XElement(ChartNs + "trendlineLbl"));
        Assert.True(OpenXmlChartSpaceCodec.TryRead(unsupported.ToString(SaveOptions.DisableFormatting), out var preserved, out _, out var unsupportedEditable));
        Assert.False(unsupportedEditable);
        Assert.Empty(preserved.Series[0].Trendlines);

        var topologyChanged = imported.Clone();
        topologyChanged.Series[0].Trendlines.RemoveAt(0);
        var error = Assert.Throws<CodecException>(() => OpenXmlChartSpaceCodec.Patch(editableDocument, topologyChanged, "trendline_topology_changed", "Test chart"));
        Assert.Equal("trendline_topology_changed", error.Code);
    }

    [Fact]
    public void TrendlineValidationRejectsInvalidTypeSpecificState()
    {
        var wrongOrder = Chart();
        wrongOrder.Series[0].Trendlines[0].PolynomialOrder = 3;
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([wrongOrder], "summary")).Code);

        var shortAverage = Chart();
        shortAverage.Series[0].Values.Clear();
        shortAverage.Series[0].Values.Add([1, 2]);
        shortAverage.Categories.Clear();
        shortAverage.Categories.Add(["A", "B"]);
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([shortAverage], "summary")).Code);

        var fractionalForecast = Chart();
        fractionalForecast.Series[0].Trendlines[1].Forward = 0.25;
        Assert.Equal("invalid_spreadsheet_chart", Assert.Throws<CodecException>(() => XlsxChartCodec.Validate([fractionalForecast], "summary")).Code);
    }

    private static SpreadsheetChartArtifact Chart()
    {
        var chart = new SpreadsheetChartArtifact
        {
            Id = "chart/trendlines",
            Name = "Trendline chart",
            Title = "Trendline evidence",
            Type = SpreadsheetChartType.Line,
            HasLegend = true,
            AbsoluteAnchor = new SpreadsheetAbsoluteAnchorArtifact { XEmu = 0, YEmu = 0, WidthEmu = 4_000_000, HeightEmu = 2_500_000 },
            XAxis = new SpreadsheetChartAxisArtifact { Title = "Period" },
            YAxis = new SpreadsheetChartAxisArtifact { Title = "Value" },
        };
        chart.Categories.Add(["A", "B", "C", "D", "E", "F"]);
        var series = new SpreadsheetChartSeriesArtifact { Name = "Revenue" };
        series.Values.Add([4, 7, 11, 18, 29, 47]);
        series.Trendlines.Add(new SpreadsheetChartTrendlineArtifact { Type = SpreadsheetChartTrendlineType.Exponential, Name = "Exponential" });
        series.Trendlines.Add(new SpreadsheetChartTrendlineArtifact
        {
            Type = SpreadsheetChartTrendlineType.Linear,
            Name = "Linear evidence",
            Forward = 0.5,
            Backward = 1,
            Intercept = 0,
            DisplayEquation = true,
            DisplayRSquared = true,
            Line = new SpreadsheetChartLineStyleArtifact
            {
                Color = new SpreadsheetColor { Rgb = "7C3AED" },
                DashStyle = SpreadsheetChartLineDashStyle.Dashed,
                WidthPoints = 1.5,
            },
        });
        series.Trendlines.Add(new SpreadsheetChartTrendlineArtifact { Type = SpreadsheetChartTrendlineType.Logarithmic, Name = "Logarithmic" });
        series.Trendlines.Add(new SpreadsheetChartTrendlineArtifact { Type = SpreadsheetChartTrendlineType.MovingAverage, Name = "Moving average", Period = 3 });
        series.Trendlines.Add(new SpreadsheetChartTrendlineArtifact { Type = SpreadsheetChartTrendlineType.Polynomial, Name = "Polynomial", PolynomialOrder = 3 });
        series.Trendlines.Add(new SpreadsheetChartTrendlineArtifact { Type = SpreadsheetChartTrendlineType.Power, Name = "Power" });
        chart.Series.Add(series);
        return chart;
    }
}
