using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns one literal DrawingML gradient: direct sRGB stops plus either a linear
// angle or the canonical centered radial profile. Theme colors, transforms,
// tile rectangles, arbitrary path gradients and effect-bearing color graphs
// remain source-owned.
internal static class PptxGradientFillCodec
{
    private const int FullTurn = 360 * 60_000;

    internal static void Validate(PresentationGradientFill? source, string subject)
    {
        if (source is null) return;
        if (source.Kind is not (PresentationGradientFill.Types.Kind.Linear or PresentationGradientFill.Types.Kind.Radial))
            throw Invalid(subject, "requires a linear or radial gradient kind.");
        if (source.Stops.Count is < 2 or > 16)
            throw Invalid(subject, "requires between 2 and 16 gradient stops.");
        uint previous = 0;
        for (var index = 0; index < source.Stops.Count; index++)
        {
            var stop = source.Stops[index];
            if (stop.PositionThousandthPercent > 100_000 ||
                index > 0 && stop.PositionThousandthPercent < previous)
                throw Invalid(subject, "gradient stop positions must be ordered values from 0 to 100000.");
            previous = stop.PositionThousandthPercent;
            _ = PptxColor.Normalize(stop.ColorRgb);
            if (stop.HasOpacityThousandthPercent && stop.OpacityThousandthPercent > 100_000)
                throw Invalid(subject, "gradient stop opacity must be from 0 to 100000.");
        }
        if (source.Kind == PresentationGradientFill.Types.Kind.Radial && source.HasAngle60000)
            throw Invalid(subject, "a radial gradient cannot define a linear angle.");
        if (source.Kind == PresentationGradientFill.Types.Kind.Linear &&
            source.HasAngle60000 && source.Angle60000 is < 0 or >= FullTurn)
            throw Invalid(subject, "linear gradient angle must be normalized to one turn.");
    }

    internal static A.GradientFill Build(PresentationGradientFill source, string subject)
    {
        Validate(source, subject);
        var stops = new A.GradientStopList();
        foreach (var semantic in source.Stops)
        {
            var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(semantic.ColorRgb) };
            if (semantic.HasOpacityThousandthPercent)
                color.Append(new A.Alpha { Val = checked((int)semantic.OpacityThousandthPercent) });
            stops.Append(new A.GradientStop(color) { Position = checked((int)semantic.PositionThousandthPercent) });
        }
        var output = new A.GradientFill(stops);
        if (source.Kind == PresentationGradientFill.Types.Kind.Linear)
        {
            output.Append(new A.LinearGradientFill
            {
                Angle = source.HasAngle60000 ? source.Angle60000 : 0,
                Scaled = false,
            });
        }
        else
        {
            output.Append(new A.PathGradientFill(
                new A.FillToRectangle { Left = 50_000, Top = 50_000, Right = 50_000, Bottom = 50_000 })
            {
                Path = A.PathShadeValues.Circle,
            });
        }
        return output;
    }

    internal static bool TryRead(A.GradientFill? source, out PresentationGradientFill semantic)
    {
        semantic = new PresentationGradientFill();
        if (source is null || source.GetAttributes().Count != 0 || source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not A.GradientStopList stopList ||
            stopList.GetAttributes().Count != 0 || stopList.ChildElements.Count is < 2 or > 16)
            return false;
        uint previous = 0;
        foreach (var nativeStop in stopList.ChildElements)
        {
            if (nativeStop is not A.GradientStop stop ||
                stop.GetAttributes().Any(attribute => attribute.LocalName != "pos") ||
                stop.Position?.Value is not int position || position is < 0 or > 100_000 ||
                semantic.Stops.Count > 0 && checked((uint)position) < previous ||
                stop.ChildElements.Count != 1 || stop.FirstChild is not A.RgbColorModelHex color ||
                color.Val?.Value is not { Length: 6 } rgb || !rgb.All(Uri.IsHexDigit) ||
                color.GetAttributes().Any(attribute => attribute.LocalName != "val"))
                return false;
            var alphas = color.Elements<A.Alpha>().ToArray();
            if (color.ChildElements.Any(child => child is not A.Alpha) || alphas.Length > 1)
                return false;
            var semanticStop = new PresentationGradientStop
            {
                PositionThousandthPercent = checked((uint)position),
                ColorRgb = rgb.ToUpperInvariant(),
            };
            if (alphas.SingleOrDefault() is { } alpha)
            {
                if (alpha.ChildElements.Count != 0 ||
                    alpha.GetAttributes().Any(attribute => attribute.LocalName != "val") ||
                    alpha.Val?.Value is not int opacity || opacity is < 0 or > 100_000)
                    return false;
                semanticStop.OpacityThousandthPercent = checked((uint)opacity);
            }
            semantic.Stops.Add(semanticStop);
            previous = semanticStop.PositionThousandthPercent;
        }
        switch (source.ChildElements[1])
        {
            case A.LinearGradientFill linear
                when linear.GetAttributes().All(attribute => attribute.LocalName is "ang" or "scaled") &&
                     linear.ChildElements.Count == 0 &&
                     linear.Angle?.Value is int angle && angle is >= 0 and < FullTurn &&
                     linear.Scaled?.Value == false:
                semantic.Kind = PresentationGradientFill.Types.Kind.Linear;
                semantic.Angle60000 = angle;
                return true;
            case A.PathGradientFill path
                when path.GetAttributes().All(attribute => attribute.LocalName == "path") &&
                     path.Path?.Value == A.PathShadeValues.Circle &&
                     path.ChildElements.Count == 1 &&
                     path.FillToRectangle is { } rectangle &&
                     rectangle.ChildElements.Count == 0 &&
                     rectangle.GetAttributes().All(attribute => attribute.LocalName is "l" or "t" or "r" or "b") &&
                     rectangle.Left?.Value == 50_000 && rectangle.Top?.Value == 50_000 &&
                     rectangle.Right?.Value == 50_000 && rectangle.Bottom?.Value == 50_000:
                semantic.Kind = PresentationGradientFill.Types.Kind.Radial;
                return true;
            default:
                semantic = new PresentationGradientFill();
                return false;
        }
    }

    private static CodecException Invalid(string subject, string detail) =>
        new("invalid_presentation_gradient", $"{subject} {detail}");
}
