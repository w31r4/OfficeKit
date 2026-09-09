using System.IO.Compression;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("shape", "columnDirection")]
    [InlineData("shape", "verticalText")]
    [InlineData("shape", "wrap")]
    [InlineData("shape", "horizontalOverflow")]
    [InlineData("shape", "verticalOverflow")]
    [InlineData("shape", "verticalAlignment")]
    [InlineData("text", "columnDirection")]
    [InlineData("text", "verticalText")]
    [InlineData("text", "wrap")]
    [InlineData("text", "horizontalOverflow")]
    [InlineData("text", "verticalOverflow")]
    [InlineData("text", "verticalAlignment")]
    [InlineData("master", "columnDirection")]
    [InlineData("master", "verticalText")]
    [InlineData("master", "wrap")]
    [InlineData("master", "horizontalOverflow")]
    [InlineData("master", "verticalOverflow")]
    [InlineData("master", "verticalAlignment")]
    [InlineData("layout", "columnDirection")]
    [InlineData("layout", "verticalText")]
    [InlineData("layout", "wrap")]
    [InlineData("layout", "horizontalOverflow")]
    [InlineData("layout", "verticalOverflow")]
    [InlineData("layout", "verticalAlignment")]
    [InlineData("table", "columnDirection")]
    [InlineData("table", "verticalText")]
    [InlineData("table", "wrap")]
    [InlineData("table", "horizontalOverflow")]
    [InlineData("table", "verticalOverflow")]
    [InlineData("table", "verticalAlignment")]
    public void EnumBodyPropertySourceRemovalAndRestorationPreserveOtherState(string kind, string field)
    {
        var program = Program(kind); Style(program, kind)[field] = BodyPropertyValues(field).Last();
        if (field == "horizontalOverflow") Style(program, kind)["verticalOverflow"] = "ellipsis";
        if (field == "verticalOverflow") Style(program, kind)["horizontalOverflow"] = "clip";
        if (field == "verticalAlignment") Style(program, kind)["anchorCenter"] = true;
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        Assert.Equal(BodyPropertyValues(field).Last(), Style(request, kind)[field]!.GetValue<string>());
        Assert.Equal(source, Compile(request, source).File.ToByteArray());
        AssertBodyPropertyValue(source, kind, field, BodyPropertyValues(field).Last());
        Style(request, kind).Remove(field);
        var deleted = Compile(request, source).File.ToByteArray();
        AssertBodyPropertyValue(deleted, kind, field, null);
        Assert.False(Style(Project(deleted), kind).ContainsKey(field));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
        foreach (var value in BodyPropertyValues(field))
        {
            var restore = Project(deleted);
            Style(restore, kind)[field] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            AssertBodyPropertyValue(restored, kind, field, value);
            Assert.Equal(value, Style(Project(restored), kind)[field]!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove(field);
            AssertBodyPropertyValue(Compile(removeAgain, restored).File.ToByteArray(), kind, field, null);
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape", "columnDirection")]
    [InlineData("shape", "verticalText")]
    [InlineData("shape", "wrap")]
    [InlineData("shape", "horizontalOverflow")]
    [InlineData("shape", "verticalOverflow")]
    [InlineData("shape", "verticalAlignment")]
    [InlineData("table", "columnDirection")]
    [InlineData("table", "verticalText")]
    [InlineData("table", "wrap")]
    [InlineData("table", "horizontalOverflow")]
    [InlineData("table", "verticalOverflow")]
    [InlineData("table", "verticalAlignment")]
    public void EnumBodyPropertyOnlyAndCombinedRemovableStylesPreservePresence(string kind, string field)
    {
        foreach (var withOtherProperties in new[] { false, true })
        {
            var program = Program(kind); var style = Style(program, kind);
            style.Clear(); style[field] = BodyPropertyValues(field).Last();
            if (withOtherProperties)
            {
                style["upright"] = false; style["rotation"] = 12;
                foreach (var sibling in new[] { "columnDirection", "verticalText", "wrap", "horizontalOverflow", "verticalOverflow", "verticalAlignment" }.Where(key => key != field))
                    style[sibling] = BodyPropertyValues(sibling).First();
            }
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            AssertBodyPropertyValue(deleted, kind, field, null);
            Assert.Null(Body(deleted, kind).UpRight);
            Assert.Null(Body(deleted, kind).Rotation);
            Assert.Null(Body(deleted, kind).Vertical);
            Assert.Null(Body(deleted, kind).RightToLeftColumns);
            Assert.Null(Body(deleted, kind).Wrap);
            Assert.Null(Body(deleted, kind).HorizontalOverflow);
            Assert.Null(Body(deleted, kind).VerticalOverflow);
            Assert.Null(Body(deleted, kind).Anchor);
            if (!withOtherProperties) AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
            var restore = Project(deleted);
            var element = restore["pages"]![0]!["elements"]![0]!;
            if (kind == "shape") element["textStyle"] = new JsonObject { [field] = BodyPropertyValues(field).First() };
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = new JsonObject { [field] = BodyPropertyValues(field).First() },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            AssertBodyPropertyValue(restored, kind, field, BodyPropertyValues(field).First());
            Assert.Equal(BodyPropertyValues(field).First(), Style(Project(restored), kind)[field]!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
        }
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("text")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void ColumnGapSourceRemovalAndRestorationPreserveOtherState(string kind)
    {
        var program = Program(kind); var style = Style(program, kind);
        style["columnGap"] = 12.5; style["columns"] = 3; style["columnDirection"] = "right-to-left";
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        Assert.Equal(12.5, Style(request, kind)["columnGap"]!.GetValue<double>());
        Assert.Equal(source, Compile(request, source).File.ToByteArray());
        Assert.Equal(158750, Body(source, kind).ColumnSpacing!.Value);
        Style(request, kind).Remove("columnGap");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(Body(deleted, kind).ColumnSpacing);
        Assert.False(Style(Project(deleted), kind).ContainsKey("columnGap"));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, "columnGap");
        foreach (var value in new[] { 0d, 12.5d, 10000d })
        {
            var restore = Project(deleted);
            Style(restore, kind)["columnGap"] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(checked((int)(value * 12700)), Body(restored, kind).ColumnSpacing!.Value);
            Assert.Equal(value, Style(Project(restored), kind)["columnGap"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "columnGap");
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove("columnGap");
            Assert.Null(Body(Compile(removeAgain, restored).File.ToByteArray(), kind).ColumnSpacing);
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void ColumnGapOnlyAndCombinedRemovableStylesPreservePresence(string kind)
    {
        foreach (var withOtherProperties in new[] { false, true })
        {
            var program = Program(kind); var style = Style(program, kind);
            style.Clear(); style["columnGap"] = 12.5;
            if (withOtherProperties)
            {
                style["upright"] = false; style["rotation"] = 12;
                foreach (var field in new[] { "columnDirection", "verticalText", "wrap", "horizontalOverflow", "verticalOverflow", "verticalAlignment" })
                    style[field] = BodyPropertyValues(field).First();
            }
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Null(Body(deleted, kind).ColumnSpacing);
            Assert.Null(Body(deleted, kind).UpRight);
            if (withOtherProperties)
            {
                Assert.Null(Body(deleted, kind).Rotation);
                foreach (var field in new[] { "columnDirection", "verticalText", "wrap", "horizontalOverflow", "verticalOverflow", "verticalAlignment" })
                    AssertBodyPropertyValue(deleted, kind, field, null);
            }
            if (!withOtherProperties) AssertOnlyBodyPropertyChanged(source, deleted, kind, "columnGap");
            var restore = Project(deleted);
            var element = restore["pages"]![0]!["elements"]![0]!;
            if (kind == "shape") element["textStyle"] = new JsonObject { ["columnGap"] = 0 };
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = new JsonObject { ["columnGap"] = 0 },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(0, Body(restored, kind).ColumnSpacing!.Value);
            Assert.Equal(0, Style(Project(restored), kind)["columnGap"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "columnGap");
        }
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("text")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void RotationSourceRemovalAndRestorationPreserveOtherState(string kind)
    {
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(Program(kind)).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        Assert.Equal(12, Style(request, kind)["rotation"]!.GetValue<double>());
        Assert.Equal(source, Compile(request, source).File.ToByteArray());
        Assert.Equal(12 * 60000, Body(source, kind).Rotation!.Value);
        Style(request, kind).Remove("rotation");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(Body(deleted, kind).Rotation);
        Assert.False(Style(Project(deleted), kind).ContainsKey("rotation"));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, "rotation");
        foreach (var value in new[] { 0, -24 })
        {
            var restore = Project(deleted);
            Style(restore, kind)["rotation"] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(value * 60000, Body(restored, kind).Rotation!.Value);
            Assert.Equal(value, Style(Project(restored), kind)["rotation"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "rotation");
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove("rotation");
            Assert.Null(Body(Compile(removeAgain, restored).File.ToByteArray(), kind).Rotation);
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void RotationOnlyAndCombinedRemovableStylesPreservePresence(string kind)
    {
        foreach (var withOtherProperties in new[] { false, true })
        {
            var program = Program(kind); var style = Style(program, kind);
            style.Clear(); style["rotation"] = 12;
            if (withOtherProperties) style["upright"] = false;
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Null(Body(deleted, kind).Rotation);
            Assert.Null(Body(deleted, kind).UpRight);
            if (!withOtherProperties) AssertOnlyBodyPropertyChanged(source, deleted, kind, "rotation");
            var restore = Project(deleted);
            var element = restore["pages"]![0]!["elements"]![0]!;
            if (kind == "shape") element["textStyle"] = new JsonObject { ["rotation"] = 0 };
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = new JsonObject { ["rotation"] = 0 },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(0, Body(restored, kind).Rotation!.Value);
            Assert.Equal(0, Style(Project(restored), kind)["rotation"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "rotation");
        }
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void UprightOnlyStyleCanBeRemovedWithExistingAuthority(string kind)
    {
        var program = Program(kind);
        Style(program, kind).Clear(); Style(program, kind)["upright"] = true;
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var request = Project(source);
        Assert.Equal(new[] { "upright" }, Style(request, kind).Select(p => p.Key));
        JsonObject StyleOwner(JsonObject input) => (kind == "shape" ? input["pages"]![0]!["elements"]![0] :
            input["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
        StyleOwner(request).Remove(kind == "shape" ? "textStyle" : "style");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(Body(deleted, kind).UpRight);
        var fresh = Project(deleted);
        if (kind == "shape") Assert.False(StyleOwner(fresh).ContainsKey("textStyle"));
        else
        {
            var cell = fresh["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!;
            Assert.Equal("Retain this text", cell["text"]!.GetValue<string>());
            cell["text"] = new JsonObject { ["style"] = new JsonObject { ["upright"] = false },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }) };
            var restored = Compile(fresh, deleted).File.ToByteArray();
            Assert.False(Body(restored, kind).UpRight!.Value);
            Assert.False(Style(Project(restored), kind)["upright"]!.GetValue<bool>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind);
        }
        AssertOnlyBodyPropertyChanged(source, deleted, kind);

        var denied = Project(source);
        StyleOwner(denied).Remove(kind == "shape" ? "textStyle" : "style");
        var capabilities = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray();
        capabilities.Remove(capabilities.Single(c => c!["operation"]!.GetValue<string>() == (kind == "shape" ? "setTextBodyStyle" : "setTableCellStyle")));
        Assert.Empty(Compile(denied, source, success: false).File);

        var otherSource = PptxCodecTests.RemoveEmbeddedPpj(Compile(Program(kind)).File.ToByteArray());
        var otherEdit = Project(otherSource);
        StyleOwner(otherEdit).Remove(kind == "shape" ? "textStyle" : "style");
        Assert.Empty(Compile(otherEdit, otherSource, success: false).File);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("text")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void UprightSourceRemovalAndRestorationPreserveOtherState(string kind)
    {
        var program = Program(kind);
        var authored = Compile(program);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(Style(projected, kind)["upright"]!.GetValue<bool>());
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        var baseline = Body(source, kind);
        Assert.True(baseline.UpRight!.Value);
        var request = Project(source);
        Style(request, kind).Remove("upright");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(Body(deleted, kind).UpRight);
        Assert.False(Style(Project(deleted), kind).ContainsKey("upright"));
        AssertOnlyBodyPropertyChanged(source, deleted, kind);
        foreach (var value in new[] { false, true })
        {
            var restore = Project(deleted);
            Style(restore, kind)["upright"] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(value, Body(restored, kind).UpRight!.Value);
            Assert.Equal(value, Style(Project(restored), kind)["upright"]!.GetValue<bool>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind);
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove("upright");
            Assert.Null(Body(Compile(removeAgain, restored).File.ToByteArray(), kind).UpRight);
        }
        Assert.Equal(original, source);
    }

    private static string[] BodyPropertyValues(string field) => field switch
    {
        "verticalText" => ["horizontal", "vertical", "vertical270"],
        "columnDirection" => ["left-to-right", "right-to-left"],
        "wrap" => ["none", "square"],
        "horizontalOverflow" => ["overflow", "clip"],
        "verticalOverflow" => ["overflow", "ellipsis", "clip"],
        "verticalAlignment" => ["top", "middle", "bottom"],
        _ => throw new ArgumentException(field),
    };

    private static void AssertBodyPropertyValue(byte[] bytes, string kind, string field, string? value)
    {
        var body = Body(bytes, kind);
        if (field == "verticalAlignment")
        {
            if (value is null) Assert.Null(body.Anchor);
            else Assert.Equal(value switch
            {
                "top" => A.TextAnchoringTypeValues.Top,
                "middle" => A.TextAnchoringTypeValues.Center,
                "bottom" => A.TextAnchoringTypeValues.Bottom,
                _ => throw new ArgumentException(value),
            }, body.Anchor!.Value);
        }
        else if (field == "verticalOverflow")
        {
            if (value is null) Assert.Null(body.VerticalOverflow);
            else Assert.Equal(value switch
            {
                "overflow" => A.TextVerticalOverflowValues.Overflow,
                "ellipsis" => A.TextVerticalOverflowValues.Ellipsis,
                "clip" => A.TextVerticalOverflowValues.Clip,
                _ => throw new ArgumentException(value),
            }, body.VerticalOverflow!.Value);
        }
        else if (field == "horizontalOverflow")
        {
            if (value is null) Assert.Null(body.HorizontalOverflow);
            else Assert.Equal(value == "clip" ? A.TextHorizontalOverflowValues.Clip : A.TextHorizontalOverflowValues.Overflow, body.HorizontalOverflow!.Value);
        }
        else if (field == "wrap")
        {
            if (value is null) Assert.Null(body.Wrap);
            else Assert.Equal(value == "square" ? A.TextWrappingValues.Square : A.TextWrappingValues.None, body.Wrap!.Value);
        }
        else if (field == "columnDirection")
        {
            if (value is null) Assert.Null(body.RightToLeftColumns);
            else Assert.Equal(value == "right-to-left", body.RightToLeftColumns!.Value);
        }
        else
        {
            if (value is null) Assert.Null(body.Vertical);
            else Assert.Equal(value switch
            {
                "horizontal" => A.TextVerticalValues.Horizontal,
                "vertical" => A.TextVerticalValues.Vertical,
                "vertical270" => A.TextVerticalValues.Vertical270,
                _ => throw new ArgumentException(value),
            }, body.Vertical!.Value);
        }
    }

    private static JsonObject Program(string kind)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "package.json"))) directory = directory.Parent;
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(directory!.FullName, "examples/ppj/minimum.ppj")))!.AsObject();
        var element = JsonNode.Parse("""
            {"id":"owner","type":"shape","frame":{"x":30,"y":40,"width":200,"height":100},
             "geometry":{"kind":"preset","preset":"rect"},"text":"Retain this text",
             "textStyle":{"upright":true,"rotation":12,"wrap":"square","margins":{"left":7}}}
            """)!.AsObject();
        if (kind == "table")
        {
            element = JsonNode.Parse("""
                {"id":"owner","type":"table","frame":{"x":30,"y":40,"width":200,"height":100},
                 "columns":[{"id":"column","width":200}],"rows":[{"id":"row","height":100,"cells":[{"id":"cell",
                 "text":{"style":{"upright":true,"rotation":12,"wrap":"square","margins":{"left":7}},
                 "paragraphs":[{"runs":[{"text":"Retain this text"}]}]}}]}]}
                """)!.AsObject();
        }
        else if (kind == "text")
        {
            element["type"] = kind;
            element["style"] = element["textStyle"]!.DeepClone(); element.Remove("textStyle"); element.Remove("geometry");
        }
        program["pages"]![0]!["elements"] = new JsonArray(element);
        if (kind is "master" or "layout")
        {
            program["design"]!["masters"] = JsonNode.Parse("""[{"id":"master","name":"Master","placeholders":[]}]""");
            program["design"]!["layouts"] = JsonNode.Parse("""[{"id":"layout","name":"Layout","master":"master","layoutType":"blank","placeholders":[]}]""");
            var placeholder = new JsonObject { ["id"] = "placeholder", ["name"] = "Owner", ["placeholderType"] = "body", ["index"] = 0,
                ["frame"] = element["frame"]!.DeepClone(), ["text"] = element["text"]!.DeepClone(), ["style"] = element["textStyle"]!.DeepClone() };
            program["design"]![kind == "master" ? "masters" : "layouts"]![0]!["placeholders"] = new JsonArray(placeholder);
            program["pages"]![0]!["layout"] = "layout";
            program["pages"]![0]!["elements"] = new JsonArray();
        }
        return program;
    }

    private static JsonObject Style(JsonObject program, string kind)
    {
        if (kind is "master" or "layout") return program["design"]![kind == "master" ? "masters" : "layouts"]![0]!["placeholders"]![0]!["style"]!.AsObject();
        var element = program["pages"]![0]!["elements"]![0]!;
        return (kind == "table" ? element["rows"]![0]!["cells"]![0]!["text"]!["style"] : element[kind == "shape" ? "textStyle" : "style"])!.AsObject();
    }
    private static CodecResponse Compile(JsonObject program, byte[]? source = null, bool success = true) => Invoke(new CodecRequest
    {
        Operation = CodecOperation.CompilePpjToPptx, File = source is null ? ByteString.Empty : ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()) },
    }, success);
    private static JsonObject Project(byte[] source) => JsonNode.Parse(Invoke(new CodecRequest
    {
        Operation = CodecOperation.ProjectPptxToPpj, File = ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { SourceUri = "source.pptx" },
    }).PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    private static CodecResponse Invoke(CodecRequest request, bool success = true)
    {
        request.ProtocolVersion = CodecProtocol.ProtocolVersion; request.Family = ArtifactFamily.Presentation;
        var bytes = request.ToByteArray(); var response = PpjCodecProtocol.InvokeResponse(ref bytes, null);
        Assert.True(response.Ok == success, string.Join("\n", response.Diagnostics.Select(d => d.Code + ": " + d.Message)));
        return response;
    }
    private static OpenXmlPartRootElement Owner(PresentationDocument document, string kind) => kind switch
    {
        "master" => document.PresentationPart!.SlideMasterParts.Single().SlideMaster,
        "layout" => document.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout,
        _ => document.PresentationPart!.SlideParts.Single().Slide,
    };
    private static A.BodyProperties Body(byte[] bytes, string kind)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        return (A.BodyProperties)Assert.Single(Owner(document, kind).Descendants<A.BodyProperties>()).CloneNode(true);
    }
    private static void AssertOnlyBodyPropertyChanged(byte[] before, byte[] after, string kind, string field = "upright")
    {
        using var left = PresentationDocument.Open(new MemoryStream(before), false);
        using var right = PresentationDocument.Open(new MemoryStream(after), false);
        var oldSlide = Owner(left, kind);
        var newSlide = Owner(right, kind);
        if (field == "columnGap")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).ColumnSpacing = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).ColumnSpacing = null;
        }
        else if (field == "verticalAlignment")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Anchor = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Anchor = null;
        }
        else if (field == "verticalOverflow")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).VerticalOverflow = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).VerticalOverflow = null;
        }
        else if (field == "horizontalOverflow")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).HorizontalOverflow = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).HorizontalOverflow = null;
        }
        else if (field == "wrap")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Wrap = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Wrap = null;
        }
        else if (field == "verticalText")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Vertical = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Vertical = null;
        }
        else if (field == "columnDirection")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).RightToLeftColumns = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).RightToLeftColumns = null;
        }
        else if (field == "rotation")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Rotation = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Rotation = null;
        }
        else
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).UpRight = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).UpRight = null;
        }
        XElement Content(string xml)
        {
            var node = XElement.Parse(xml);
            foreach (var attribute in node.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).ToArray()) attribute.Remove();
            return node;
        }
        // The table writer may repeat namespace declarations. Compare expanded
        // XML names/values while keeping all non-namespace attributes/content.
        Assert.True(XNode.DeepEquals(Content(oldSlide.OuterXml), Content(newSlide.OuterXml)));
        using var oldZip = new ZipArchive(new MemoryStream(before));
        using var newZip = new ZipArchive(new MemoryStream(after));
        Assert.Equal(oldZip.Entries.Select(e => e.FullName).Order(), newZip.Entries.Select(e => e.FullName).Order());
        var ownerPart = kind switch { "master" => "ppt/slideMasters/slideMaster1.xml", "layout" => "ppt/slideLayouts/slideLayout1.xml", _ => "ppt/slides/slide1.xml" };
        foreach (var entry in oldZip.Entries.Where(e => e.FullName != ownerPart))
        {
            using var a = entry.Open(); using var b = newZip.GetEntry(entry.FullName)!.Open();
            using var aa = new MemoryStream(); using var bb = new MemoryStream(); a.CopyTo(aa); b.CopyTo(bb);
            Assert.Equal(aa.ToArray(), bb.ToArray());
        }
    }
}
