using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjAuthoredMediaPlaybackTriggerAuthorsStartConditionAndPreservesClickDefault()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(fixtureDirectory, "evidence-ledger-canonical.ppj")))!.AsObject();
        var mediaBytes = Convert.FromHexString("000000186674797069736F6D0000020069736F6D6D703431");
        var mediaSha256 = Convert.ToHexString(SHA256.HashData(mediaBytes)).ToLowerInvariant();
        var posterBytes = File.ReadAllBytes(Path.Combine(fixtureDirectory, "ppj-assets", "evidence-mark.svg"));
        var posterSha256 = Convert.ToHexString(SHA256.HashData(posterBytes)).ToLowerInvariant();

        var page = program["pages"]![0]!.AsObject();
        page.Remove("animations");
        page.Remove("timing");
        page.Remove("transition");
        page.Remove("notes");
        page["elements"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "media-trigger",
                ["name"] = "Media trigger",
                ["role"] = "authored playback trigger",
                ["type"] = "media",
                ["frame"] = new JsonObject { ["x"] = 80, ["y"] = 80, ["width"] = 320, ["height"] = 180 },
                ["accessibility"] = new JsonObject
                {
                    ["decorative"] = false,
                    ["description"] = "A video used to verify authored playback timing.",
                },
                ["mediaType"] = "video",
                ["asset"] = "evidence-video",
                ["posterAsset"] = "evidence-mark",
                ["playback"] = new JsonObject { ["trigger"] = "onSlideStart" },
            },
        };
        page["readingOrder"] = new JsonArray("media-trigger");
        program["pages"] = new JsonArray(page.DeepClone());
        program["assets"] = new JsonArray
        {
            program["assets"]![0]!.DeepClone(),
            new JsonObject
            {
                ["id"] = "evidence-video",
                ["uri"] = "ppj-assets/evidence-video.mp4",
                ["mimeType"] = "video/mp4",
                ["sha256"] = mediaSha256,
                ["rights"] = new JsonObject { ["status"] = "internal" },
                ["accessibility"] = new JsonObject
                {
                    ["decorative"] = false,
                    ["description"] = "Synthetic authored video for the timing contract.",
                },
            },
        };
        program["components"] = new JsonArray();
        program["sections"] = new JsonArray();
        program["customShows"] = new JsonArray();
        program["comments"] = new JsonArray();

        CodecResponse Compile(JsonObject candidate) => Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(candidate.ToJsonString()),
                Assets =
                {
                    new Asset
                    {
                        Id = "evidence-mark",
                        FileName = "evidence-mark.svg",
                        ContentType = "image/svg+xml",
                        Data = ByteString.CopyFrom(posterBytes),
                        Sha256 = posterSha256,
                    },
                    new Asset
                    {
                        Id = "evidence-video",
                        FileName = "evidence-video.mp4",
                        ContentType = "video/mp4",
                        Data = ByteString.CopyFrom(mediaBytes),
                        Sha256 = mediaSha256,
                    },
                },
            },
        });

        var programBytes = Encoding.UTF8.GetBytes(program.ToJsonString());
        var validation = PpjProgramValidator.Validate(programBytes);
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics));

        var authored = Compile(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        Assert.Equal(programBytes, ZipBytes(authored.File.ToByteArray(), "officeKit/program.ppj"));
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var timing = package.PresentationPart!.SlideParts.Single().Slide!.Timing!.OuterXml;
            Assert.Contains("<p:video", timing, StringComparison.Ordinal);
            Assert.Contains("delay=\"0\"", timing, StringComparison.Ordinal);
        }

        var clickProgram = program.DeepClone().AsObject();
        var clickMedia = clickProgram["pages"]![0]!["elements"]![0]!.AsObject();
        clickMedia["playback"] = new JsonObject { ["trigger"] = "onClick" };
        var clickStart = Compile(clickProgram);
        Assert.True(clickStart.Ok, Diagnostics(clickStart));
        using (var stream = new MemoryStream(clickStart.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var timing = package.PresentationPart!.SlideParts.Single().Slide!.Timing!.OuterXml;
            Assert.Contains("delay=\"indefinite\"", timing, StringComparison.Ordinal);
        }

        var omittedProgram = program.DeepClone().AsObject();
        omittedProgram["pages"]![0]!["elements"]![0]!.AsObject().Remove("playback");
        var omittedStart = Compile(omittedProgram);
        Assert.True(omittedStart.Ok, Diagnostics(omittedStart));
        using (var stream = new MemoryStream(omittedStart.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var timing = package.PresentationPart!.SlideParts.Single().Slide!.Timing!.OuterXml;
            Assert.Contains("delay=\"indefinite\"", timing, StringComparison.Ordinal);
        }

        var invalidProgram = program.DeepClone().AsObject();
        invalidProgram["pages"]![0]!["elements"]![0]!["playback"]!["trigger"] = "afterVideo";
        var invalid = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidProgram.ToJsonString()));
        Assert.False(invalid.IsValid);
    }
}
