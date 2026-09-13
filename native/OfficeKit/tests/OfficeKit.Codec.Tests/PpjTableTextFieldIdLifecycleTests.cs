using System.IO.Compression;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Fact]
    public void StaticTableTextFieldIdLifecyclePatchesOnlyFieldIdToken()
    {
        const string originalId = "{11111111-2222-4333-8444-555555555555}";
        const string replacementId = "{22222222-3333-4444-8555-666666666666}";
        var program = Program("table");
        program["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"runs":[{"field":{"id":"__ID__","type":"customStatic","text":"cached"}}]}]}
            """.Replace("__ID__", originalId, StringComparison.Ordinal));
        var authored = Compile(program).File.ToByteArray();
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored);
        var originalSource = source.ToArray();

        var projected = Project(source);
        var table = projected["pages"]![0]!["elements"]![0]!.AsObject();
        var cell = table["rows"]![0]!["cells"]![0]!.AsObject();
        var field = cell["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!.AsObject();
        var idLeaf = table["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "tableTextFieldId")!.AsObject();
        Assert.Equal(originalId, idLeaf["value"]!.GetValue<string>());
        var capability = table["nativeRef"]!["capabilities"]!.AsArray()
            .Single(item => item!["operation"]!.GetValue<string>() == "setTextField");
        Assert.Contains("table.rows[].cells[].text.paragraphs[].runs[].field.id",
            capability!["fields"]!.AsArray().Select(item => item!.GetValue<string>()));

        var noOp = Compile(projected, source);
        Assert.True(noOp.Ok, string.Join("\n", noOp.Diagnostics.Select(item => item.Code + ": " + item.Message)));
        Assert.Empty(noOp.PresentationProgram.ChangedParts);
        Assert.Equal(source, noOp.File.ToByteArray());

        field["id"] = replacementId;
        var edited = Compile(projected, source);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => item.Code + ": " + item.Message)));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        Assert.Equal(originalSource, source);

        var output = edited.File.ToByteArray();
        var outputField = ReadTableField(output);
        Assert.Equal(replacementId, outputField.Id!.Value);
        Assert.Equal("customStatic", outputField.Type!.Value);
        Assert.Equal("cached", outputField.Text!.Text);
        using (var document = PresentationDocument.Open(new MemoryStream(output), false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        foreach (var path in ZipPaths(source).Union(ZipPaths(output), StringComparer.Ordinal))
            if (!string.Equals(path, "ppt/slides/slide1.xml", StringComparison.Ordinal))
                Assert.Equal(ZipBytes(source, path), ZipBytes(output, path));

        var reprojected = Project(output);
        var reprojectedField = reprojected["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!;
        Assert.Equal(replacementId, reprojectedField!["id"]!.GetValue<string>());
        Assert.Equal("customStatic", reprojectedField["type"]!.GetValue<string>());
        Assert.Equal(replacementId, reprojected["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "tableTextFieldId")!["value"]!.GetValue<string>());

        var stale = Project(source);
        stale["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["id"] = replacementId;
        stale["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "tableTextFieldId")!["expectedHash"] = new string('0', 64);
        Assert.False(Compile(stale, source, success: false).Ok);

        var invalid = Project(source);
        invalid["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["id"] = "not-a-guid";
        Assert.False(Compile(invalid, source, success: false).Ok);

        var automatic = Project(source);
        automatic["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["type"] = "slidenum";
        automatic["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["id"] = replacementId;
        Assert.False(Compile(automatic, source, success: false).Ok);

        var combined = Project(source);
        var combinedField = combined["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!.AsObject();
        combinedField["id"] = replacementId;
        combinedField["text"] = "changed";
        Assert.False(Compile(combined, source, success: false).Ok);
        Assert.Equal(originalSource, source);
    }
}
