using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Bounded DrawingML custom paths used by source-built presentation templates.
// Coordinates and arc values may reference one ordered adjustment/guide graph;
// formula parsing and evaluation stay in PptxCustomGeometryFormulaCodec. Shape-
// local text bounds retain the exact four private scaling guides emitted below.
// Non-empty handles, connection sites, formula text rectangles, 3D, and relative
// lighten/darken path fill remain opaque and fail closed.
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

    private sealed class Profile
    {
        internal required A.PathList Paths { get; init; }
        internal required PptxCustomGeometryFormulaCodec.Graph Formulas { get; init; }
        internal PresentationCustomGeometryTextRectangle? TextRectangle { get; init; }
    }

    internal static bool Supports(A.CustomGeometry? geometry, long widthEmu, long heightEmu)
    {
        return TryProfile(geometry, widthEmu, heightEmu, out _);
    }

    internal static void Read(A.CustomGeometry? geometry, long widthEmu, long heightEmu, PresentationShape target)
    {
        if (!TryProfile(geometry, widthEmu, heightEmu, out var profile)) return;
        target.CustomAdjustments.Add(profile.Formulas.Adjustments);
        target.CustomGuides.Add(profile.Formulas.Guides);
        target.TextRectangle = profile.TextRectangle;
        foreach (var nativePath in profile.Paths.Elements<A.Path>())
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
            target.CustomPaths.Add(path);
        }
    }

    private static bool TryProfile(A.CustomGeometry? geometry, long widthEmu, long heightEmu, out Profile profile)
    {
        profile = null!;
        if (geometry is null || geometry.HasAttributes || geometry.ChildElements.Count < 2 ||
            geometry.ChildElements[0] is not A.AdjustValueList adjustments)
            return false;
        var index = 1;
        A.ShapeGuideList? guideList = null;
        if (index < geometry.ChildElements.Count && geometry.ChildElements[index] is A.ShapeGuideList sourceGuides)
        {
            if (sourceGuides.HasAttributes) return false;
            guideList = sourceGuides;
            index++;
        }
        if (index < geometry.ChildElements.Count && geometry.ChildElements[index] is A.AdjustHandleList handles)
        {
            if (handles.HasAttributes || handles.ChildElements.Count != 0) return false;
            index++;
        }
        if (index < geometry.ChildElements.Count && geometry.ChildElements[index] is A.ConnectionSiteList connections)
        {
            if (connections.HasAttributes || connections.ChildElements.Count != 0) return false;
            index++;
        }
        A.Rectangle? nativeRectangle = null;
        if (index < geometry.ChildElements.Count && geometry.ChildElements[index] is A.Rectangle rectangle)
        {
            nativeRectangle = rectangle;
            index++;
        }
        if (index != geometry.ChildElements.Count - 1 || geometry.ChildElements[index] is not A.PathList pathList || pathList.HasAttributes)
            return false;
        var allGuides = guideList?.Elements<A.ShapeGuide>().ToArray() ?? [];
        if (guideList is not null && guideList.ChildElements.Count != allGuides.Length) return false;
        if (!TryReadTextRectangle(nativeRectangle, allGuides, widthEmu, heightEmu, out var textRectangle, out var userGuideCount))
            return false;
        if (!PptxCustomGeometryFormulaCodec.TryRead(adjustments, allGuides.Take(userGuideCount), widthEmu, heightEmu, out var formulas))
            return false;
        var paths = pathList.Elements<A.Path>().ToArray();
        if (paths.Length is < 1 or > MaxPaths || pathList.ChildElements.Count != paths.Length) return false;
        var commandCount = 0;
        if (!paths.All(path => Supports(path, formulas, ref commandCount))) return false;
        profile = new Profile { Paths = pathList, Formulas = formulas, TextRectangle = textRectangle };
        return true;
    }

    internal static void Validate(PresentationShape shape, string shapeId)
    {
        if (shape.Geometry != "custom")
        {
            if (shape.CustomPaths.Count > 0 || shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 || shape.TextRectangle is not null)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has custom geometry data without custom geometry.");
            return;
        }
        if (shape.CustomPaths.Count is < 1 or > MaxPaths)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} custom geometry must contain 1 through {MaxPaths} paths.");
        var formulas = PptxCustomGeometryFormulaCodec.Validate(shape, shapeId);
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
                Validate(command, shapeId, hasCurrentPoint, formulas);
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
        var adjustments = new A.AdjustValueList(shape.CustomAdjustments.Select(PptxCustomGeometryFormulaCodec.Write));
        var geometry = new A.CustomGeometry(adjustments);
        A.ShapeGuideList? guides = null;
        if (shape.CustomGuides.Count > 0 || shape.TextRectangle is not null)
        {
            guides = new A.ShapeGuideList(shape.CustomGuides.Select(PptxCustomGeometryFormulaCodec.Write));
            if (shape.TextRectangle is not null)
                guides.Append(TextRectangleGuides(shape.TextRectangle, widthEmu, heightEmu));
            geometry.Append(guides);
        }
        if (shape.TextRectangle is not null)
            geometry.Append(TextRectangle());
        geometry.Append(paths);
        return geometry;
    }

    private static bool Supports(A.Path path, PptxCustomGeometryFormulaCodec.Graph formulas, ref int commandCount)
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
                A.MoveTo move => SupportsPointContainer(move, move.Point, 1, formulas),
                A.LineTo line => SupportsPointContainer(line, line.Point, 1, formulas),
                A.QuadraticBezierCurveTo quadratic => SupportsPointSequence(quadratic, 2, formulas),
                A.CubicBezierCurveTo cubic => SupportsPointSequence(cubic, 3, formulas),
                A.ArcTo arc => hasCurrentPoint && SupportsArc(arc, formulas),
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
        A.Rectangle? nativeRectangle,
        IReadOnlyList<A.ShapeGuide> guides,
        long widthEmu,
        long heightEmu,
        out PresentationCustomGeometryTextRectangle? rectangle,
        out int userGuideCount)
    {
        rectangle = null;
        userGuideCount = guides.Count;
        if (nativeRectangle is null) return true;
        if (SupportsLiteralTextRectangle(nativeRectangle))
        {
            rectangle = new PresentationCustomGeometryTextRectangle
            {
                LeftEmu = ParseCoordinate(nativeRectangle.Left!.Value!),
                TopEmu = ParseCoordinate(nativeRectangle.Top!.Value!),
                RightEmu = ParseCoordinate(nativeRectangle.Right!.Value!),
                BottomEmu = ParseCoordinate(nativeRectangle.Bottom!.Value!),
            };
            return true;
        }
        if (guides.Count < 4 || !HasOnlyAttributes(nativeRectangle, "l", "t", "r", "b") || !HasNoInnerXml(nativeRectangle) ||
            nativeRectangle.Left?.Value != TextLeftGuide || nativeRectangle.Top?.Value != TextTopGuide ||
            nativeRectangle.Right?.Value != TextRightGuide || nativeRectangle.Bottom?.Value != TextBottomGuide)
            return false;
        userGuideCount = guides.Count - 4;
        if (!TryScaledGuide(guides[userGuideCount], TextLeftGuide, "w", widthEmu, out var left) ||
            !TryScaledGuide(guides[userGuideCount + 1], TextTopGuide, "h", heightEmu, out var top) ||
            !TryScaledGuide(guides[userGuideCount + 2], TextRightGuide, "w", widthEmu, out var right) ||
            !TryScaledGuide(guides[userGuideCount + 3], TextBottomGuide, "h", heightEmu, out var bottom) ||
            left >= right || top >= bottom)
            return false;
        rectangle = new PresentationCustomGeometryTextRectangle
        {
            LeftEmu = left,
            TopEmu = top,
            RightEmu = right,
            BottomEmu = bottom,
        };
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

    private static bool SupportsPointContainer(OpenXmlCompositeElement container, A.Point? point, int childCount, PptxCustomGeometryFormulaCodec.Graph formulas) =>
        !container.HasAttributes && container.ChildElements.Count == childCount && point is not null && SupportsPoint(point, formulas);

    private static bool SupportsPointSequence(OpenXmlCompositeElement container, int childCount, PptxCustomGeometryFormulaCodec.Graph formulas) =>
        !container.HasAttributes && container.ChildElements.Count == childCount &&
        container.ChildElements.All(child => child is A.Point point && SupportsPoint(point, formulas));

    private static bool SupportsPoint(A.Point point, PptxCustomGeometryFormulaCodec.Graph formulas) =>
        HasNoInnerXml(point) && HasOnlyAttributes(point, "x", "y") &&
        TryValue(point.X?.Value, formulas, out _) && TryValue(point.Y?.Value, formulas, out _);

    private static bool SupportsArc(A.ArcTo arc, PptxCustomGeometryFormulaCodec.Graph formulas) =>
        HasNoInnerXml(arc) && HasOnlyAttributes(arc, "wR", "hR", "stAng", "swAng") &&
        TryValue(arc.WidthRadius?.Value, formulas, out var widthRadius) && widthRadius > 0 &&
        TryValue(arc.HeightRadius?.Value, formulas, out var heightRadius) && heightRadius > 0 &&
        TryValue(arc.StartAngle?.Value, formulas, out _) &&
        TryValue(arc.SwingAngle?.Value, formulas, out var sweepAngle) &&
        sweepAngle != 0 && Math.Abs(sweepAngle) <= FullTurnAngle;

    private static PresentationCustomGeometryPoint ReadPoint(A.Point point)
    {
        var result = new PresentationCustomGeometryPoint();
        if (TryCoordinate(point.X?.Value, out var x)) result.X = x;
        else result.XReference = point.X!.Value!;
        if (TryCoordinate(point.Y?.Value, out var y)) result.Y = y;
        else result.YReference = point.Y!.Value!;
        return result;
    }

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

    private static PresentationCustomGeometryCommand ReadArc(A.ArcTo source)
    {
        var arc = new PresentationCustomGeometryArc();
        if (TryCoordinate(source.WidthRadius?.Value, out var widthRadius)) arc.WidthRadius = widthRadius;
        else arc.WidthRadiusReference = source.WidthRadius!.Value!;
        if (TryCoordinate(source.HeightRadius?.Value, out var heightRadius)) arc.HeightRadius = heightRadius;
        else arc.HeightRadiusReference = source.HeightRadius!.Value!;
        if (TryAngle(source.StartAngle?.Value, out var startAngle)) arc.StartAngle = startAngle;
        else arc.StartAngleReference = source.StartAngle!.Value!;
        if (TryAngle(source.SwingAngle?.Value, out var sweepAngle)) arc.SweepAngle = sweepAngle;
        else arc.SweepAngleReference = source.SwingAngle!.Value!;
        return new PresentationCustomGeometryCommand { ArcTo = arc };
    }

    private static A.Point Point(PresentationCustomGeometryPoint source) => new()
    {
        X = source.HasXReference ? source.XReference : source.X.ToString(CultureInfo.InvariantCulture),
        Y = source.HasYReference ? source.YReference : source.Y.ToString(CultureInfo.InvariantCulture),
    };

    private static A.ArcTo Arc(PresentationCustomGeometryArc source) => new()
    {
        WidthRadius = source.HasWidthRadiusReference ? source.WidthRadiusReference : source.WidthRadius.ToString(CultureInfo.InvariantCulture),
        HeightRadius = source.HasHeightRadiusReference ? source.HeightRadiusReference : source.HeightRadius.ToString(CultureInfo.InvariantCulture),
        StartAngle = source.HasStartAngleReference ? source.StartAngleReference : source.StartAngle.ToString(CultureInfo.InvariantCulture),
        SwingAngle = source.HasSweepAngleReference ? source.SweepAngleReference : source.SweepAngle.ToString(CultureInfo.InvariantCulture),
    };

    private static IEnumerable<A.ShapeGuide> TextRectangleGuides(PresentationCustomGeometryTextRectangle source, long widthEmu, long heightEmu) => [
        ScaledGuide(TextLeftGuide, source.LeftEmu, "w", widthEmu),
        ScaledGuide(TextTopGuide, source.TopEmu, "h", heightEmu),
        ScaledGuide(TextRightGuide, source.RightEmu, "w", widthEmu),
        ScaledGuide(TextBottomGuide, source.BottomEmu, "h", heightEmu)];

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

    private static void Validate(
        PresentationCustomGeometryCommand command,
        string shapeId,
        bool hasCurrentPoint,
        PptxCustomGeometryFormulaCodec.Graph formulas)
    {
        switch (command.CommandCase)
        {
            case PresentationCustomGeometryCommand.CommandOneofCase.MoveTo:
                Validate(command.MoveTo, shapeId, formulas);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.LineTo:
                Validate(command.LineTo, shapeId, formulas);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo:
                if (command.QuadraticBezierTo.Control is null || command.QuadraticBezierTo.End is null)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an incomplete quadratic Bézier command.");
                Validate(command.QuadraticBezierTo.Control, shapeId, formulas);
                Validate(command.QuadraticBezierTo.End, shapeId, formulas);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo:
                if (command.CubicBezierTo.Control1 is null || command.CubicBezierTo.Control2 is null || command.CubicBezierTo.End is null)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an incomplete cubic Bézier command.");
                Validate(command.CubicBezierTo.Control1, shapeId, formulas);
                Validate(command.CubicBezierTo.Control2, shapeId, formulas);
                Validate(command.CubicBezierTo.End, shapeId, formulas);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.ArcTo:
                if (!hasCurrentPoint)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an arc command without an established current point.");
                if (command.ArcTo is null ||
                    !TryWireValue(command.ArcTo.HasWidthRadiusReference, command.ArcTo.WidthRadiusReference, command.ArcTo.WidthRadius, formulas, out var widthRadius) ||
                    !TryWireValue(command.ArcTo.HasHeightRadiusReference, command.ArcTo.HeightRadiusReference, command.ArcTo.HeightRadius, formulas, out var heightRadius) ||
                    !TryWireValue(command.ArcTo.HasStartAngleReference, command.ArcTo.StartAngleReference, command.ArcTo.StartAngle, formulas, out _) ||
                    !TryWireValue(command.ArcTo.HasSweepAngleReference, command.ArcTo.SweepAngleReference, command.ArcTo.SweepAngle, formulas, out var sweepAngle) ||
                    widthRadius <= 0 || heightRadius <= 0 || sweepAngle == 0 || Math.Abs(sweepAngle) > FullTurnAngle)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid bounded custom arc command or formula reference.");
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.Close:
                if (!command.Close) throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid close command.");
                break;
            default:
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} contains an empty custom geometry command.");
        }
    }

    private static void Validate(PresentationCustomGeometryPoint? point, string shapeId, PptxCustomGeometryFormulaCodec.Graph formulas)
    {
        if (point is null ||
            !TryWireValue(point.HasXReference, point.XReference, point.X, formulas, out _) ||
            !TryWireValue(point.HasYReference, point.YReference, point.Y, formulas, out _))
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid custom path point or formula reference.");
    }

    private static bool TryWireValue(
        bool hasReference,
        string reference,
        long literal,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        out double value)
    {
        value = 0;
        if (hasReference && literal != 0) return false;
        return formulas.TryResolve(hasReference ? reference : null, literal, out value);
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

    private static bool TryValue(string? source, PptxCustomGeometryFormulaCodec.Graph formulas, out double value)
    {
        if (TryCoordinate(source, out var coordinate))
        {
            value = coordinate;
            return true;
        }
        return formulas.TryResolveReference(source, out value);
    }

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
