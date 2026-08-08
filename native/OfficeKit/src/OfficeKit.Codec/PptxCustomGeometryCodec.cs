using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Bounded literal DrawingML custom paths used by source-built presentation
// templates. Shape-local text bounds accept native numeric rectangles and the
// exact four scaling guides emitted below for cross-host rendering. Every other
// formula geometry, guide, handle, connection site, and relative lighten/darken
// path fill stays outside this slice.
internal static class PptxCustomGeometryCodec
{
    private const int MaxPaths = 64;
    private const int MaxCommands = 16_384;
    private const long MaxCoordinate = int.MaxValue;
    private const int FullTurnAngle = 21_600_000;
    private const string TextLeftGuide = "officeKitTextLeft";
    private const string TextTopGuide = "officeKitTextTop";
    private const string TextRightGuide = "officeKitTextRight";
    private const string TextBottomGuide = "officeKitTextBottom";

    internal static bool Supports(A.CustomGeometry? geometry, long widthEmu, long heightEmu)
    {
        if (geometry is null || geometry.HasAttributes || geometry.ChildElements.Count is < 2 or > 4 ||
            geometry.ChildElements[0] is not A.AdjustValueList adjustValues || adjustValues.HasChildren || adjustValues.HasAttributes ||
            geometry.ChildElements[^1] is not A.PathList pathList || pathList.HasAttributes)
            return false;
        var hasRectangle = TryReadTextRectangle(geometry, widthEmu, heightEmu, out _);
        if (geometry.ChildElements.Count == 2 ? hasRectangle : !hasRectangle) return false;
        var paths = pathList.Elements<A.Path>().ToArray();
        if (paths.Length is < 1 or > MaxPaths || pathList.ChildElements.Count != paths.Length) return false;
        var commandCount = 0;
        return paths.All(path => Supports(path, ref commandCount));
    }

    internal static IEnumerable<PresentationCustomGeometryPath> Read(A.CustomGeometry? geometry, long widthEmu, long heightEmu)
    {
        if (!Supports(geometry, widthEmu, heightEmu)) yield break;
        foreach (var nativePath in geometry!.GetFirstChild<A.PathList>()!.Elements<A.Path>())
        {
            var path = new PresentationCustomGeometryPath
            {
                Width = checked((long)nativePath.Width!.Value),
                Height = checked((long)nativePath.Height!.Value),
            };
            if (nativePath.Fill?.HasValue == true)
                path.FillMode = nativePath.Fill.Value == A.PathFillModeValues.None
                    ? PresentationCustomGeometryPath.Types.FillMode.None
                    : PresentationCustomGeometryPath.Types.FillMode.Normal;
            if (nativePath.Stroke?.HasValue == true) path.Stroke = nativePath.Stroke.Value;
            if (nativePath.ExtrusionOk?.HasValue == true) path.ExtrusionAllowed = nativePath.ExtrusionOk.Value;
            foreach (var nativeCommand in nativePath.ChildElements)
            {
                var command = nativeCommand switch
                {
                    A.MoveTo move => new PresentationCustomGeometryCommand { MoveTo = ReadPoint(move.Point!) },
                    A.LineTo line => new PresentationCustomGeometryCommand { LineTo = ReadPoint(line.Point!) },
                    A.QuadraticBezierCurveTo quadratic => ReadQuadratic(quadratic),
                    A.CubicBezierCurveTo cubic => ReadCubic(cubic),
                    A.ArcTo arc => ReadArc(arc),
                    A.CloseShapePath => new PresentationCustomGeometryCommand { Close = true },
                    _ => throw new InvalidOperationException("Unsupported custom geometry command passed the recognition gate."),
                };
                path.Commands.Add(command);
            }
            yield return path;
        }
    }

    internal static PresentationCustomGeometryTextRectangle? ReadTextRectangle(A.CustomGeometry? geometry, long widthEmu, long heightEmu)
    {
        return Supports(geometry, widthEmu, heightEmu) && TryReadTextRectangle(geometry!, widthEmu, heightEmu, out var rectangle)
            ? rectangle
            : null;
    }

    internal static void Validate(PresentationShape shape, string shapeId)
    {
        if (shape.Geometry != "custom")
        {
            if (shape.CustomPaths.Count > 0 || shape.TextRectangle is not null)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has custom geometry data without custom geometry.");
            return;
        }
        if (shape.CustomPaths.Count is < 1 or > MaxPaths)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} custom geometry must contain 1 through {MaxPaths} paths.");
        Validate(shape.TextRectangle, shapeId);
        var commandCount = 0;
        foreach (var path in shape.CustomPaths)
        {
            if (path.Width is <= 0 or > MaxCoordinate || path.Height is <= 0 or > MaxCoordinate || path.Commands.Count == 0)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid custom path extent or empty command list.");
            if (path.FillMode is not (PresentationCustomGeometryPath.Types.FillMode.Unspecified or
                PresentationCustomGeometryPath.Types.FillMode.Normal or PresentationCustomGeometryPath.Types.FillMode.None))
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an unsupported custom path fill mode.");
            commandCount += path.Commands.Count;
            if (commandCount > MaxCommands)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation shape {shapeId} custom geometry exceeds the {MaxCommands}-command budget.");
            var hasCurrentPoint = false;
            var hasSubpathStart = false;
            foreach (var command in path.Commands)
            {
                Validate(command, shapeId, hasCurrentPoint);
                switch (command.CommandCase)
                {
                    case PresentationCustomGeometryCommand.CommandOneofCase.MoveTo:
                        hasCurrentPoint = true;
                        hasSubpathStart = true;
                        break;
                    case PresentationCustomGeometryCommand.CommandOneofCase.LineTo:
                    case PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo:
                    case PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo:
                        hasCurrentPoint = true;
                        break;
                    case PresentationCustomGeometryCommand.CommandOneofCase.Close:
                        hasCurrentPoint = hasSubpathStart;
                        break;
                }
            }
        }
    }

    internal static void Apply(P.ShapeProperties properties, PresentationShape shape)
    {
        if (shape.Geometry != "custom")
        {
            properties.GetFirstChild<A.CustomGeometry>()?.Remove();
            var preset = properties.GetFirstChild<A.PresetGeometry>();
            if (preset is null)
            {
                preset = new A.PresetGeometry(new A.AdjustValueList());
                var presetTransform = properties.GetFirstChild<A.Transform2D>();
                if (presetTransform is null) properties.PrependChild(preset);
                else properties.InsertAfter(preset, presetTransform);
            }
            preset.Preset = shape.Geometry switch
            {
                "ellipse" => A.ShapeTypeValues.Ellipse,
                "roundRect" => A.ShapeTypeValues.RoundRectangle,
                _ => A.ShapeTypeValues.Rectangle,
            };
            return;
        }
        properties.GetFirstChild<A.PresetGeometry>()?.Remove();
        properties.GetFirstChild<A.CustomGeometry>()?.Remove();
        var transform = properties.GetFirstChild<A.Transform2D>();
        var widthEmu = transform?.Extents?.Cx?.Value ?? shape.WidthEmu;
        var heightEmu = transform?.Extents?.Cy?.Value ?? shape.HeightEmu;
        OpenXmlElement geometry = Build(shape, widthEmu, heightEmu);
        if (transform is null) properties.PrependChild(geometry);
        else properties.InsertAfter(geometry, transform);
    }

    private static A.CustomGeometry Build(PresentationShape shape, long widthEmu, long heightEmu)
    {
        var paths = new A.PathList();
        foreach (var source in shape.CustomPaths)
        {
            var path = new A.Path { Width = source.Width, Height = source.Height };
            if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.Normal) path.Fill = A.PathFillModeValues.Norm;
            else if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.None) path.Fill = A.PathFillModeValues.None;
            if (source.HasStroke) path.Stroke = source.Stroke;
            if (source.HasExtrusionAllowed) path.ExtrusionOk = source.ExtrusionAllowed;
            foreach (var command in source.Commands)
            {
                path.Append(command.CommandCase switch
                {
                    PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => new A.MoveTo(Point(command.MoveTo)),
                    PresentationCustomGeometryCommand.CommandOneofCase.LineTo => new A.LineTo(Point(command.LineTo)),
                    PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo => new A.QuadraticBezierCurveTo(
                        Point(command.QuadraticBezierTo.Control),
                        Point(command.QuadraticBezierTo.End)),
                    PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo => new A.CubicBezierCurveTo(
                        Point(command.CubicBezierTo.Control1),
                        Point(command.CubicBezierTo.Control2),
                        Point(command.CubicBezierTo.End)),
                    PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => Arc(command.ArcTo),
                    PresentationCustomGeometryCommand.CommandOneofCase.Close => new A.CloseShapePath(),
                    _ => throw new CodecException("invalid_presentation_geometry", "Presentation custom geometry contains an empty command."),
                });
            }
            paths.Append(path);
        }
        var geometry = new A.CustomGeometry(new A.AdjustValueList());
        if (shape.TextRectangle is not null)
        {
            geometry.Append(TextRectangleGuides(shape.TextRectangle, widthEmu, heightEmu));
            geometry.Append(TextRectangle());
        }
        geometry.Append(paths);
        return geometry;
    }

    private static bool Supports(A.Path path, ref int commandCount)
    {
        if (path.Width?.Value is not { } width || width is 0 or > MaxCoordinate ||
            path.Height?.Value is not { } height || height is 0 or > MaxCoordinate ||
            !HasOnlyAttributes(path, "w", "h", "fill", "stroke", "extrusionOk") || !SupportsPathProperties(path) ||
            path.ChildElements.Count == 0)
            return false;
        commandCount += path.ChildElements.Count;
        if (commandCount > MaxCommands) return false;
        var hasCurrentPoint = false;
        var hasSubpathStart = false;
        foreach (var command in path.ChildElements)
        {
            var supported = command switch
            {
                A.MoveTo move => SupportsPointContainer(move, move.Point, 1),
                A.LineTo line => SupportsPointContainer(line, line.Point, 1),
                A.QuadraticBezierCurveTo quadratic => SupportsPointSequence(quadratic, 2),
                A.CubicBezierCurveTo cubic => SupportsPointSequence(cubic, 3),
                A.ArcTo arc => hasCurrentPoint && SupportsArc(arc),
                A.CloseShapePath close => !close.HasAttributes && HasNoInnerXml(close),
                _ => false,
            };
            if (!supported) return false;
            switch (command)
            {
                case A.MoveTo:
                    hasCurrentPoint = true;
                    hasSubpathStart = true;
                    break;
                case A.LineTo:
                case A.QuadraticBezierCurveTo:
                case A.CubicBezierCurveTo:
                    hasCurrentPoint = true;
                    break;
                case A.CloseShapePath:
                    hasCurrentPoint = hasSubpathStart;
                    break;
            }
        }
        return true;
    }

    private static bool SupportsLiteralTextRectangle(A.Rectangle rectangle) =>
        HasNoInnerXml(rectangle) && HasOnlyAttributes(rectangle, "l", "t", "r", "b") &&
        TryCoordinate(rectangle.Left?.Value, out var left) &&
        TryCoordinate(rectangle.Top?.Value, out var top) &&
        TryCoordinate(rectangle.Right?.Value, out var right) &&
        TryCoordinate(rectangle.Bottom?.Value, out var bottom) &&
        left < right && top < bottom;

    private static bool TryReadTextRectangle(
        A.CustomGeometry geometry,
        long widthEmu,
        long heightEmu,
        out PresentationCustomGeometryTextRectangle rectangle)
    {
        rectangle = new PresentationCustomGeometryTextRectangle();
        if (geometry.GetFirstChild<A.Rectangle>() is not { } nativeRectangle) return false;
        if (geometry.GetFirstChild<A.ShapeGuideList>() is not { } guides)
        {
            if (geometry.ChildElements.Count != 3 || geometry.ChildElements[1] != nativeRectangle || !SupportsLiteralTextRectangle(nativeRectangle)) return false;
            rectangle.LeftEmu = ParseCoordinate(nativeRectangle.Left!.Value!);
            rectangle.TopEmu = ParseCoordinate(nativeRectangle.Top!.Value!);
            rectangle.RightEmu = ParseCoordinate(nativeRectangle.Right!.Value!);
            rectangle.BottomEmu = ParseCoordinate(nativeRectangle.Bottom!.Value!);
            return true;
        }
        if (geometry.ChildElements.Count != 4 || geometry.ChildElements[1] != guides || geometry.ChildElements[2] != nativeRectangle ||
            guides.HasAttributes || guides.ChildElements.Count != 4 ||
            !HasOnlyAttributes(nativeRectangle, "l", "t", "r", "b") || !HasNoInnerXml(nativeRectangle) ||
            nativeRectangle.Left?.Value != TextLeftGuide || nativeRectangle.Top?.Value != TextTopGuide ||
            nativeRectangle.Right?.Value != TextRightGuide || nativeRectangle.Bottom?.Value != TextBottomGuide)
            return false;
        var source = guides.Elements<A.ShapeGuide>().ToArray();
        if (source.Length != 4 ||
            !TryScaledGuide(source[0], TextLeftGuide, "w", widthEmu, out var left) ||
            !TryScaledGuide(source[1], TextTopGuide, "h", heightEmu, out var top) ||
            !TryScaledGuide(source[2], TextRightGuide, "w", widthEmu, out var right) ||
            !TryScaledGuide(source[3], TextBottomGuide, "h", heightEmu, out var bottom) ||
            left >= right || top >= bottom)
            return false;
        rectangle.LeftEmu = left;
        rectangle.TopEmu = top;
        rectangle.RightEmu = right;
        rectangle.BottomEmu = bottom;
        return true;
    }

    private static bool TryScaledGuide(A.ShapeGuide guide, string name, string axis, long extentEmu, out long coordinate)
    {
        coordinate = 0;
        if (!HasNoInnerXml(guide) || !HasOnlyAttributes(guide, "name", "fmla") || guide.Name?.Value != name || extentEmu <= 0) return false;
        var tokens = (guide.Formula?.Value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return tokens.Length == 4 && tokens[0] == "*/" && tokens[2] == axis &&
            TryCoordinate(tokens[1], out coordinate) &&
            long.TryParse(tokens[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out var denominator) && denominator == extentEmu;
    }

    private static bool SupportsPathProperties(A.Path path)
    {
        if (path.Fill is not null)
        {
            if (path.Fill.HasValue != true) return false;
            var fill = path.Fill.Value;
            if (fill != A.PathFillModeValues.Norm && fill != A.PathFillModeValues.None) return false;
        }
        return path.Stroke is not { HasValue: false } && path.ExtrusionOk is not { HasValue: false };
    }

    private static bool SupportsPointContainer(OpenXmlCompositeElement container, A.Point? point, int childCount) =>
        !container.HasAttributes && container.ChildElements.Count == childCount && point is not null && SupportsPoint(point);

    private static bool SupportsPointSequence(OpenXmlCompositeElement container, int childCount) =>
        !container.HasAttributes && container.ChildElements.Count == childCount &&
        container.ChildElements.All(child => child is A.Point point && SupportsPoint(point));

    private static bool SupportsPoint(A.Point point) =>
        HasNoInnerXml(point) && HasOnlyAttributes(point, "x", "y") &&
        TryCoordinate(point.X?.Value, out _) && TryCoordinate(point.Y?.Value, out _);

    private static bool SupportsArc(A.ArcTo arc) =>
        HasNoInnerXml(arc) && HasOnlyAttributes(arc, "wR", "hR", "stAng", "swAng") &&
        TryCoordinate(arc.WidthRadius?.Value, out var widthRadius) && widthRadius > 0 &&
        TryCoordinate(arc.HeightRadius?.Value, out var heightRadius) && heightRadius > 0 &&
        TryAngle(arc.StartAngle?.Value, out _) &&
        TryAngle(arc.SwingAngle?.Value, out var sweepAngle) &&
        sweepAngle != 0 && Math.Abs((long)sweepAngle) <= FullTurnAngle;

    private static PresentationCustomGeometryPoint ReadPoint(A.Point point) => new()
    {
        X = ParseCoordinate(point.X!.Value!),
        Y = ParseCoordinate(point.Y!.Value!),
    };

    private static PresentationCustomGeometryCommand ReadCubic(A.CubicBezierCurveTo source)
    {
        var points = source.Elements<A.Point>().ToArray();
        return new PresentationCustomGeometryCommand
        {
            CubicBezierTo = new PresentationCustomGeometryCubicBezier
            {
                Control1 = ReadPoint(points[0]),
                Control2 = ReadPoint(points[1]),
                End = ReadPoint(points[2]),
            },
        };
    }

    private static PresentationCustomGeometryCommand ReadQuadratic(A.QuadraticBezierCurveTo source)
    {
        var points = source.Elements<A.Point>().ToArray();
        return new PresentationCustomGeometryCommand
        {
            QuadraticBezierTo = new PresentationCustomGeometryQuadraticBezier
            {
                Control = ReadPoint(points[0]),
                End = ReadPoint(points[1]),
            },
        };
    }

    private static PresentationCustomGeometryCommand ReadArc(A.ArcTo source) => new()
    {
        ArcTo = new PresentationCustomGeometryArc
        {
            WidthRadius = ParseCoordinate(source.WidthRadius!.Value!),
            HeightRadius = ParseCoordinate(source.HeightRadius!.Value!),
            StartAngle = ParseAngle(source.StartAngle!.Value!),
            SweepAngle = ParseAngle(source.SwingAngle!.Value!),
        },
    };

    private static A.Point Point(PresentationCustomGeometryPoint source) => new()
    {
        X = source.X.ToString(CultureInfo.InvariantCulture),
        Y = source.Y.ToString(CultureInfo.InvariantCulture),
    };

    private static A.ArcTo Arc(PresentationCustomGeometryArc source) => new()
    {
        WidthRadius = source.WidthRadius.ToString(CultureInfo.InvariantCulture),
        HeightRadius = source.HeightRadius.ToString(CultureInfo.InvariantCulture),
        StartAngle = source.StartAngle.ToString(CultureInfo.InvariantCulture),
        SwingAngle = source.SweepAngle.ToString(CultureInfo.InvariantCulture),
    };

    private static A.ShapeGuideList TextRectangleGuides(PresentationCustomGeometryTextRectangle source, long widthEmu, long heightEmu) => new(
        ScaledGuide(TextLeftGuide, source.LeftEmu, "w", widthEmu),
        ScaledGuide(TextTopGuide, source.TopEmu, "h", heightEmu),
        ScaledGuide(TextRightGuide, source.RightEmu, "w", widthEmu),
        ScaledGuide(TextBottomGuide, source.BottomEmu, "h", heightEmu));

    private static A.ShapeGuide ScaledGuide(string name, long coordinate, string axis, long extentEmu) => new()
    {
        Name = name,
        Formula = $"*/ {coordinate.ToString(CultureInfo.InvariantCulture)} {axis} {extentEmu.ToString(CultureInfo.InvariantCulture)}",
    };

    private static A.Rectangle TextRectangle() => new()
    {
        Left = TextLeftGuide,
        Top = TextTopGuide,
        Right = TextRightGuide,
        Bottom = TextBottomGuide,
    };

    private static void Validate(PresentationCustomGeometryCommand command, string shapeId, bool hasCurrentPoint)
    {
        switch (command.CommandCase)
        {
            case PresentationCustomGeometryCommand.CommandOneofCase.MoveTo:
                Validate(command.MoveTo, shapeId);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.LineTo:
                Validate(command.LineTo, shapeId);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo:
                if (command.QuadraticBezierTo.Control is null || command.QuadraticBezierTo.End is null)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an incomplete quadratic Bézier command.");
                Validate(command.QuadraticBezierTo.Control, shapeId);
                Validate(command.QuadraticBezierTo.End, shapeId);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo:
                if (command.CubicBezierTo.Control1 is null || command.CubicBezierTo.Control2 is null || command.CubicBezierTo.End is null)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an incomplete cubic Bézier command.");
                Validate(command.CubicBezierTo.Control1, shapeId);
                Validate(command.CubicBezierTo.Control2, shapeId);
                Validate(command.CubicBezierTo.End, shapeId);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.ArcTo:
                if (!hasCurrentPoint)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an arc command without an established current point.");
                if (command.ArcTo is null || command.ArcTo.WidthRadius is <= 0 or > MaxCoordinate || command.ArcTo.HeightRadius is <= 0 or > MaxCoordinate ||
                    command.ArcTo.SweepAngle == 0 || Math.Abs((long)command.ArcTo.SweepAngle) > FullTurnAngle)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid bounded literal arc command.");
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.Close:
                if (!command.Close) throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid close command.");
                break;
            default:
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} contains an empty custom geometry command.");
        }
    }

    private static void Validate(PresentationCustomGeometryPoint? point, string shapeId)
    {
        if (point is null || point.X < -MaxCoordinate || point.X > MaxCoordinate || point.Y < -MaxCoordinate || point.Y > MaxCoordinate)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has a custom path point outside the signed 32-bit coordinate range.");
    }

    private static void Validate(PresentationCustomGeometryTextRectangle? rectangle, string shapeId)
    {
        if (rectangle is null) return;
        if (rectangle.LeftEmu < -MaxCoordinate || rectangle.LeftEmu > MaxCoordinate ||
            rectangle.TopEmu < -MaxCoordinate || rectangle.TopEmu > MaxCoordinate ||
            rectangle.RightEmu < -MaxCoordinate || rectangle.RightEmu > MaxCoordinate ||
            rectangle.BottomEmu < -MaxCoordinate || rectangle.BottomEmu > MaxCoordinate ||
            rectangle.LeftEmu >= rectangle.RightEmu || rectangle.TopEmu >= rectangle.BottomEmu)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid literal custom-geometry text rectangle.");
    }

    private static bool TryCoordinate(string? value, out long coordinate) =>
        long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out coordinate) &&
        coordinate >= -MaxCoordinate && coordinate <= MaxCoordinate;

    private static bool TryAngle(string? value, out int angle) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out angle);

    private static long ParseCoordinate(string value) => long.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);

    private static int ParseAngle(string value) => int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);

    // OpenXmlLeafElement.HasChildren is always false even when malformed source
    // XML is retained in its shadow element. InnerXml is the actual lexical gate.
    private static bool HasNoInnerXml(OpenXmlElement element) => string.IsNullOrEmpty(element.InnerXml);

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }
}
