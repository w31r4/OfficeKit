using System.Text.Json;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Lowers PPJ's high-level image fit into the finite native image-paint state.
// Authored embedded programs retain the original fit; an ordinary PPTX import
// can recover only its executable source rectangle and stretch/tile owner.
internal static class PpjImagePaintLowering
{
    internal static PresentationImagePaint Build(
        JsonElement source,
        double frameWidth,
        double frameHeight,
        Func<string, string> resolveAsset,
        Func<string, (double Width, double Height)?> assetDimensions,
        string path)
    {
        var assetId = source.GetProperty("asset").GetString()!;
        var fit = source.TryGetProperty("fit", out var fitValue) ? fitValue.GetString()! : "stretch";
        var output = new PresentationImagePaint
        {
            AssetId = resolveAsset(assetId),
            Mode = fit == "tile"
                ? PresentationImagePaint.Types.Mode.Tile
                : PresentationImagePaint.Types.Mode.Stretch,
        };
        if (source.TryGetProperty("opacity", out var opacity))
            output.OpacityThousandthPercent = Unit(opacity.GetDouble());
        if (source.TryGetProperty("crop", out var crop))
        {
            output.Crop = Crop(crop);
            return output;
        }
        if (fit is "stretch" or "tile") return output;
        if (fit is not ("cover" or "contain"))
            throw Unsupported(path, $"unsupported image fit {fit}");
        var dimensions = assetDimensions(assetId) ??
            throw Unsupported(path, $"{fit} requires declared image dimensions");
        if (frameWidth <= 0 || frameHeight <= 0)
            throw Unsupported(path, $"{fit} requires a positive destination frame");
        var sourceAspect = dimensions.Width / dimensions.Height;
        var frameAspect = frameWidth / frameHeight;
        if (fit == "cover")
        {
            var horizontal = sourceAspect > frameAspect ? (1 - frameAspect / sourceAspect) / 2 : 0;
            var vertical = sourceAspect < frameAspect ? (1 - sourceAspect / frameAspect) / 2 : 0;
            output.Crop = Symmetric(horizontal, vertical, 1);
        }
        else
        {
            var horizontal = sourceAspect < frameAspect ? (1 - sourceAspect / frameAspect) / 2 : 0;
            var vertical = sourceAspect > frameAspect ? (1 - frameAspect / sourceAspect) / 2 : 0;
            output.Crop = Symmetric(horizontal, vertical, -1);
        }
        return output;
    }

    private static PresentationImageCrop Crop(JsonElement source) => new()
    {
        LeftThousandthPercent = Edge(source, "left"),
        TopThousandthPercent = Edge(source, "top"),
        RightThousandthPercent = Edge(source, "right"),
        BottomThousandthPercent = Edge(source, "bottom"),
    };

    private static PresentationImageCrop Symmetric(double horizontal, double vertical, int sign) => new()
    {
        LeftThousandthPercent = checked(sign * (int)Math.Round(horizontal * 100_000)),
        RightThousandthPercent = checked(sign * (int)Math.Round(horizontal * 100_000)),
        TopThousandthPercent = checked(sign * (int)Math.Round(vertical * 100_000)),
        BottomThousandthPercent = checked(sign * (int)Math.Round(vertical * 100_000)),
    };

    private static int Edge(JsonElement source, string name) =>
        source.TryGetProperty(name, out var value)
            ? checked((int)Math.Round(value.GetDouble() * 100_000))
            : 0;

    private static uint Unit(double value) => checked((uint)Math.Round(value * 100_000));

    private static CodecException Unsupported(string path, string message) =>
        new("ppj.compile.unsupported", $"PPJ {path} {message}.", path);
}
