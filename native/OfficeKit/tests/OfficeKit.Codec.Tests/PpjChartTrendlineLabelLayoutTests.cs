using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using OfficeKit.Artifact.Wire.V1;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    private const string TrendlineLayoutJson = """
        {"manual":{"target":"inner","xMode":"edge","yMode":"factor","widthMode":"factor","heightMode":"edge","x":0,"y":-0.125,"width":1.25,"height":0.4}}
        """;

    [Theory]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjTrendlineLabelLayoutAuthorsEditsRemovesAndReprojects(string chartType, int seriesIndex)
    {
        var program = TrendlineListProgram(chartType, seriesIndex);
        var label = JsonNode.Parse(TrendlineLabelJson)!.AsObject();
        label["layout"] = JsonNode.Parse(TrendlineLayoutJson);
        TrendlineListSeries(program, seriesIndex)["trendlines"]![0]!["label"] = label;
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var initial = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(c + "manualLayout").Single();
        var expected = XElement.Parse($"<c:manualLayout xmlns:c='{c}'><c:layoutTarget val='inner'/><c:xMode val='edge'/><c:yMode val='factor'/><c:wMode val='factor'/><c:hMode val='edge'/><c:x val='0'/><c:y val='-0.125'/><c:w val='1.25'/><c:h val='0.4'/></c:manualLayout>");
        Assert.Equal(expected.Elements().Select(e => (e.Name, (string?)e.Attribute("val"))), initial.Elements().Select(e => (e.Name, (string?)e.Attribute("val"))));
        var projected = ProjectTrendlineList(source);
        Assert.True(JsonNode.DeepEquals(label, TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]));
        var noOp = CompileTrendlineList(projected, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var requested in new string?[]
        {
            "{\"manual\":{\"target\":\"outer\",\"xMode\":\"factor\",\"yMode\":\"edge\",\"widthMode\":\"edge\",\"heightMode\":\"factor\",\"x\":-0.2,\"y\":0,\"width\":0.8,\"height\":0.3}}",
            "{\"manual\":{\"x\":0}}", "{\"manual\":{}}", "{}", null, TrendlineLayoutJson,
        })
        {
            var item = TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]!.AsObject();
            if (requested is null) item.Remove("layout");
            else item["layout"] = JsonNode.Parse(requested);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var before = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path)));
            var after = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var native = after.Descendants(c + "trendlineLbl").Single();
            Assert.Equal(requested is not null, native.Element(c + "layout") is not null);
            Assert.Equal(requested is not null && requested != "{}", native.Descendants(c + "manualLayout").Any());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(new C.TrendlineLabel(native.ToString(SaveOptions.DisableFormatting))));
            before.Descendants(c + "trendlineLbl").Elements(c + "layout").Remove();
            after.Descendants(c + "trendlineLbl").Elements(c + "layout").Remove();
            Assert.True(XNode.DeepEquals(before, after), "Only label layout may change inside ChartML.");
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectTrendlineList(source);
            var actual = TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]!.AsObject();
            Assert.True(JsonNode.DeepEquals(requested is null ? null : JsonNode.Parse(requested), actual["layout"]));
            var rest = actual.DeepClone().AsObject();
            rest.Remove("layout");
            Assert.True(JsonNode.DeepEquals(JsonNode.Parse(TrendlineLabelJson), rest));
        }
    }

    [Fact]
    public void PpjTrendlineLabelLayoutPreservesNativeDefaultsAndRejectsNonfiniteWire()
    {
        var program = TrendlineListProgram("line", 0);
        TrendlineListSeries(program, 0)["trendlines"]![0]!["label"] = JsonNode.Parse(TrendlineLabelJson);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        source = ReplaceZipText(source, path, text =>
        {
            var xml = XDocument.Parse(text);
            xml.Descendants(c + "trendlineLbl").Single().AddFirst(XElement.Parse($"<c:layout xmlns:c='{c}'><c:manualLayout><c:layoutTarget/><c:xMode/><c:yMode/><c:wMode/><c:hMode/></c:manualLayout></c:layout>"));
            return xml.ToString(SaveOptions.DisableFormatting);
        });
        var projected = ProjectTrendlineList(source);
        var label = TrendlineListSeries(projected, 0)["trendlines"]![0]!["label"]!;
        Assert.True(JsonNode.DeepEquals(JsonNode.Parse("""{"manual":{"target":"outer","xMode":"factor","yMode":"factor","widthMode":"factor","heightMode":"factor"}}"""), label["layout"]));
        var noOp = CompileTrendlineList(projected, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var value in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity })
            Assert.Throws<CodecException>(() => OpenXmlChartLayoutCodec.Validate(new SpreadsheetChartLayoutArtifact { Manual = new SpreadsheetChartManualLayoutArtifact { X = value } }));
    }
}
