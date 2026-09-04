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
        string path,
        Func<JsonElement, double>? resolveOpacity = null)
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
            output.OpacityThousandthPercent = Unit(resolveOpacity is null ? opacity.GetDouble() : resolveOpacity(opacity));
        if (source.TryGetProperty("crop", out var crop))
        {
            if (source.TryGetProperty("focus", out _))
                throw Unsupported(path, "explicit crop cannot be combined with a focal crop");
            output.Crop = Crop(crop);
            return output;
        }
        if (fit is "stretch" or "tile")
        {
            if (source.TryGetProperty("focus", out _))
                throw Unsupported(path, "focal crop requires cover fit");
            return output;
        }
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
            output.Crop = source.TryGetProperty("focus", out var focus)
                ? Focused(horizontal, vertical, focus, path)
                : Symmetric(horizontal, vertical, 1);
        }
        else
        {
            if (source.TryGetProperty("focus", out _))
                throw Unsupported(path, "focal crop requires cover fit");
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

    private static PresentationImageCrop Focused(
        double horizontal,
        double vertical,
        JsonElement focus,
        string path)
    {
        if (focus.ValueKind != JsonValueKind.Object ||
            !focus.TryGetProperty("x", out var x) ||
            !focus.TryGetProperty("y", out var y) ||
            !x.TryGetDouble(out var focusX) ||
            !y.TryGetDouble(out var focusY) ||
            !double.IsFinite(focusX) ||
            !double.IsFinite(focusY) ||
            focusX is < 0 or > 1 ||
            focusY is < 0 or > 1)
            throw Unsupported(path, "focal crop x and y must be finite normalized values between 0 and 1");

        var visibleWidth = 1 - 2 * horizontal;
        var visibleHeight = 1 - 2 * vertical;
        var maxLeft = Math.Max(0, 1 - visibleWidth);
        var maxTop = Math.Max(0, 1 - visibleHeight);
        var left = maxLeft == 0 ? 0 : Math.Clamp(focusX - visibleWidth / 2, 0, maxLeft);
        var top = maxTop == 0 ? 0 : Math.Clamp(focusY - visibleHeight / 2, 0, maxTop);
        return new PresentationImageCrop
        {
            LeftThousandthPercent = checked((int)Math.Round(left * 100_000)),
            RightThousandthPercent = checked((int)Math.Round((maxLeft - left) * 100_000)),
            TopThousandthPercent = checked((int)Math.Round(top * 100_000)),
            BottomThousandthPercent = checked((int)Math.Round((maxTop - top) * 100_000)),
        };
    }

    private static int Edge(JsonElement source, string name) =>
        source.TryGetProperty(name, out var value)
            ? checked((int)Math.Round(value.GetDouble() * 100_000))
            : 0;

    private static uint Unit(double value) => checked((uint)Math.Round(value * 100_000));

    private static CodecException Unsupported(string path, string message) =>
        new("ppj.compile.unsupported", $"PPJ {path} {message}.", path);
}
