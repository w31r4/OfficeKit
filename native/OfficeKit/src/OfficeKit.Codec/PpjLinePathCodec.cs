using System.Text.Json;
using System.Text.Json.Nodes;
using System.Globalization;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Shared, deliberately small linePath codec.  A line is a stroked path, not
// a connector: it has no endpoint binding or routing topology.  The native
// shape representation is reused so no new wire message is needed.  Only a
// single literal path is admitted; Kimi's authored points sugar is lowered to
// that path before writing. Guide/handle/formula graphs remain opaque.
internal static class PpjLinePathCodec
{
    private const double UnitsPerPoint = 1_000d;
    private const double EmuPerPoint = 12_700d;
    private const int KimiPointBudget = 128;

    internal static bool IsLineLike(PresentationShape shape) =>
        shape.Geometry == "custom" && shape.CustomPaths.Count == 1 &&
        shape.CustomAdjustments.Count == 0 && shape.CustomGuides.Count == 0 &&
        shape.CustomConnectionSites.Count == 0 && shape.CustomAdjustmentHandles.Count == 0 &&
        shape.TextRectangle is null &&
        shape.CustomPaths[0].Commands.Count >= 2 &&
        (shape.CustomPaths[0].FillMode is PresentationCustomGeometryPath.Types.FillMode.None or
            PresentationCustomGeometryPath.Types.FillMode.Unspecified) &&
        (!shape.CustomPaths[0].HasStroke || shape.CustomPaths[0].Stroke) &&
        shape.CustomPaths[0].Commands.All(command => command.CommandCase != PresentationCustomGeometryCommand.CommandOneofCase.Close);

    internal static JsonObject Synthetic(PresentationShape shape) =>
        new()
        {
            ["viewBox"] = new JsonObject
            {
                ["x"] = 0,
                ["y"] = 0,
                ["width"] = Math.Max(0.001, shape.WidthEmu / EmuPerPoint),
                ["height"] = Math.Max(0.001, shape.HeightEmu / EmuPerPoint),
            },
            ["commands"] = new JsonArray
            {
                new JsonObject { ["op"] = "moveTo", ["x"] = 0, ["y"] = 0 },
                new JsonObject
                {
                    ["op"] = "lineTo",
                    ["x"] = Math.Max(0.001, shape.WidthEmu / EmuPerPoint),
                    ["y"] = Math.Max(0.001, shape.HeightEmu / EmuPerPoint),
                },
            },
        };

    // Kimi PPTD's compact line form is intentionally lowered to the same
    // typed path envelope used by the native profile.  The lowering is
    // bounded: sharp/round use line segments, while smooth accepts the exact
    // 2-point line, 3-point quadratic, and 4-point cubic forms. Larger smooth
    // point sets follow Kimi's documented Bezier convention: the first and
    // last points are endpoints and every middle point is a control point of
    // one high-degree Bezier. The high-degree curve is lowered to deterministic
    // cubic segments at equal parameter intervals; the finite 128-point budget
    // keeps the authored graph bounded.
    internal static JsonElement KimiPath(JsonElement raw, double fallbackWidth, double fallbackHeight, string elementId)
    {
        if (!raw.TryGetProperty("points", out var pointsValue) || pointsValue.ValueKind != JsonValueKind.String)
            throw new CodecException("ppj.line.points", $"Line {elementId} points must be a coordinate string.");
        var pointsText = pointsValue.GetString();
        if (string.IsNullOrWhiteSpace(pointsText))
            throw new CodecException("ppj.line.points", $"Line {elementId} points cannot be empty.");

        var viewBoxWidth = fallbackWidth;
        var viewBoxHeight = fallbackHeight;
        if (raw.TryGetProperty("viewBox", out var viewBox))
        {
            if (viewBox.ValueKind != JsonValueKind.Array || viewBox.GetArrayLength() != 2)
                throw new CodecException("ppj.line.points", $"Line {elementId} points viewBox must be [width,height].");
            viewBoxWidth = Finite(viewBox[0], "viewBox[0]", elementId);
            viewBoxHeight = Finite(viewBox[1], "viewBox[1]", elementId);
        }
        if (!double.IsFinite(viewBoxWidth) || !double.IsFinite(viewBoxHeight) ||
            viewBoxWidth <= 0 || viewBoxHeight <= 0 || viewBoxWidth > 100_000 || viewBoxHeight > 100_000)
            throw new CodecException("ppj.line.points", $"Line {elementId} points viewBox must have positive bounded dimensions.");

        var points = new List<(double X, double Y)>();
        foreach (var token in pointsText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (points.Count >= KimiPointBudget)
                throw new CodecException("presentation_item_budget_exceeded", $"Line {elementId} exceeds the points budget.");
            var comma = token.IndexOf(',');
            if (comma <= 0 || comma >= token.Length - 1 || token.IndexOf(',', comma + 1) >= 0 ||
                !TryFinite(token[..comma], out var x) || !TryFinite(token[(comma + 1)..], out var y))
                throw new CodecException("ppj.line.points", $"Line {elementId} points must contain finite x,y pairs.");
            if (x < 0 || x > viewBoxWidth || y < 0 || y > viewBoxHeight)
                throw new CodecException("ppj.line.points", $"Line {elementId} point ({x.ToString(CultureInfo.InvariantCulture)},{y.ToString(CultureInfo.InvariantCulture)}) lies outside its viewBox.");
            points.Add((x, y));
        }
        if (points.Count < 2)
            throw new CodecException("ppj.line.points", $"Line {elementId} requires at least two points.");

        var curve = raw.TryGetProperty("curve", out var curveValue) && curveValue.ValueKind == JsonValueKind.String
            ? curveValue.GetString()
            : "round";
        var commands = new JsonArray { PointCommand("moveTo", points[0]) };
        switch (curve)
        {
            case "sharp":
            case "round":
                foreach (var point in points.Skip(1)) commands.Add(PointCommand("lineTo", point));
                break;
            case "smooth":
                switch (points.Count)
                {
                    case 2:
                        commands.Add(PointCommand("lineTo", points[1]));
                        break;
                    case 3:
                        commands.Add(new JsonObject
                        {
                            ["op"] = "quadraticTo",
                            ["x1"] = points[1].X,
                            ["y1"] = points[1].Y,
                            ["x"] = points[2].X,
                            ["y"] = points[2].Y,
                        });
                        break;
                    case 4:
                        commands.Add(new JsonObject
                        {
                            ["op"] = "cubicTo",
                            ["x1"] = points[1].X,
                            ["y1"] = points[1].Y,
                            ["x2"] = points[2].X,
                            ["y2"] = points[2].Y,
                            ["x"] = points[3].X,
                            ["y"] = points[3].Y,
                        });
                        break;
                    default:
                        foreach (var command in SmoothCommands(points)) commands.Add(command);
                        break;
                }
                break;
            default:
                throw new CodecException("ppj.line.points", $"Line {elementId} contains unsupported curve {curve ?? "(missing)"}.");
        }

        var path = new JsonObject
        {
            ["viewBox"] = new JsonObject
            {
                ["x"] = 0,
                ["y"] = 0,
                ["width"] = viewBoxWidth,
                ["height"] = viewBoxHeight,
            },
            ["commands"] = commands,
        };
        using var document = JsonDocument.Parse(path.ToJsonString());
        return document.RootElement.Clone();
    }

    private static IEnumerable<JsonObject> SmoothCommands(IReadOnlyList<(double X, double Y)> points)
    {
        // A degree-(n-1) Bezier has n-1 cubic pieces in the Kimi lowering.
        // Each piece uses the exact value and first derivative at its two
        // parameter boundaries. This is the Hermite-to-cubic conversion used
        // by the Kimi PPTD runtime and keeps the curve C1-continuous.
        var segmentCount = points.Count - 1;
        for (var index = 0; index < segmentCount; index++)
        {
            var startParameter = index / (double)segmentCount;
            var endParameter = (index + 1) / (double)segmentCount;
            var parameterLength = endParameter - startParameter;
            var start = BezierValue(points, startParameter);
            var end = BezierValue(points, endParameter);
            var startDerivative = BezierDerivative(points, startParameter);
            var endDerivative = BezierDerivative(points, endParameter);
            yield return new JsonObject
            {
                ["op"] = "cubicTo",
                ["x1"] = start.X + parameterLength * startDerivative.X / 3d,
                ["y1"] = start.Y + parameterLength * startDerivative.Y / 3d,
                ["x2"] = end.X - parameterLength * endDerivative.X / 3d,
                ["y2"] = end.Y - parameterLength * endDerivative.Y / 3d,
                ["x"] = end.X,
                ["y"] = end.Y,
            };
        }
    }

    private static (double X, double Y) BezierValue(
        IReadOnlyList<(double X, double Y)> points,
        double parameter)
    {
        var work = points.ToArray();
        for (var level = work.Length - 1; level > 0; level--)
            for (var index = 0; index < level; index++)
                work[index] = (
                    work[index].X + (work[index + 1].X - work[index].X) * parameter,
                    work[index].Y + (work[index + 1].Y - work[index].Y) * parameter);
        return work[0];
    }

    private static (double X, double Y) BezierDerivative(
        IReadOnlyList<(double X, double Y)> points,
        double parameter)
    {
        var degree = points.Count - 1;
        var differences = new (double X, double Y)[degree];
        for (var index = 0; index < degree; index++)
            differences[index] = (
                degree * (points[index + 1].X - points[index].X),
                degree * (points[index + 1].Y - points[index].Y));
        return BezierValue(differences, parameter);
    }

    internal static JsonObject Project(PresentationShape shape)
    {
        if (!IsLineLike(shape))
            throw new CodecException("unsupported_presentation_line", "The native path is not a literal single stroked path.");
        var source = shape.CustomPaths[0];
        var commands = new JsonArray();
        foreach (var command in source.Commands) commands.Add(ProjectCommand(command));
        return new JsonObject
        {
            ["viewBox"] = new JsonObject
            {
                ["x"] = 0,
                ["y"] = 0,
                ["width"] = source.Width / UnitsPerPoint,
                ["height"] = source.Height / UnitsPerPoint,
            },
            ["commands"] = commands,
        };
    }

    // Preserve Kimi's compact points spelling when the native path is a
    // literal moveTo followed by line segments, one quadratic/cubic segment,
    // or a chain emitted by the bounded high-degree Bezier lowering. A native
    // straight path has no separate curve marker, so the line join supplies
    // the only lossless round/sharp distinction. Arc, arbitrary multi-segment
    // Bézier, and reference-backed paths continue to use the canonical typed
    // path envelope.
    internal static JsonObject? TryProjectKimi(PresentationShape shape)
    {
        if (!IsLineLike(shape)) return null;
        var commands = shape.CustomPaths[0].Commands;
        if (commands.Count < 2 || commands[0].CommandCase != PresentationCustomGeometryCommand.CommandOneofCase.MoveTo ||
            !IsLiteral(commands[0].MoveTo))
            return null;

        var curve = "sharp";
        IEnumerable<PresentationCustomGeometryPoint> points;
        if (commands.Skip(1).All(command => command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.LineTo) &&
            commands.Skip(1).All(command => IsLiteral(command.LineTo)))
        {
            points = commands.Select(command => command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.MoveTo
                ? command.MoveTo
                : command.LineTo);
            curve = shape.LineJoin == "round" ? "round" : "sharp";
        }
        else if (commands.Count == 2 && commands[1].CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo &&
                 IsLiteral(commands[1].QuadraticBezierTo.Control) && IsLiteral(commands[1].QuadraticBezierTo.End))
        {
            points = new[]
            {
                commands[0].MoveTo,
                commands[1].QuadraticBezierTo.Control,
                commands[1].QuadraticBezierTo.End,
            };
            curve = "smooth";
        }
        else if (commands.Count == 2 && commands[1].CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo &&
                 IsLiteral(commands[1].CubicBezierTo.Control1) && IsLiteral(commands[1].CubicBezierTo.Control2) &&
                 IsLiteral(commands[1].CubicBezierTo.End))
        {
            points = new[]
            {
                commands[0].MoveTo,
                commands[1].CubicBezierTo.Control1,
                commands[1].CubicBezierTo.Control2,
                commands[1].CubicBezierTo.End,
            };
            curve = "smooth";
        }
        else if (TryProjectSmooth(commands, out var smoothPoints))
        {
            points = smoothPoints;
            curve = "smooth";
        }
        else return null;

        var pointText = string.Join(" ", points.Select(PointText));
        return new JsonObject
        {
            ["viewBox"] = new JsonArray(shape.CustomPaths[0].Width / UnitsPerPoint, shape.CustomPaths[0].Height / UnitsPerPoint),
            ["points"] = pointText,
            ["curve"] = curve,
        };
    }

    private const int MaxSmoothProjectionPoints = 24;

    private static bool TryProjectSmooth(
        IReadOnlyList<PresentationCustomGeometryCommand> commands,
        out PresentationCustomGeometryPoint[] points)
    {
        points = [];
        if (commands.Count < 3 || commands.Count - 1 > MaxSmoothProjectionPoints - 1 ||
            commands[0].CommandCase != PresentationCustomGeometryCommand.CommandOneofCase.MoveTo ||
            !IsLiteral(commands[0].MoveTo) ||
            commands.Skip(1).Any(command =>
                command.CommandCase != PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo ||
                !IsLiteral(command.CubicBezierTo.Control1) ||
                !IsLiteral(command.CubicBezierTo.Control2) ||
                !IsLiteral(command.CubicBezierTo.End)))
            return false;

        var anchors = new[] { commands[0].MoveTo }
            .Concat(commands.Skip(1).Select(command => command.CubicBezierTo.End))
            .ToArray();
        if (!TryRecoverBezierControls(anchors, commands, out points)) return false;

        var nativePoints = points.Select(point => (X: (double)point.X, Y: (double)point.Y)).ToArray();
        var segmentCount = points.Length - 1;
        for (var index = 0; index < segmentCount; index++)
        {
            var startParameter = index / (double)segmentCount;
            var endParameter = (index + 1) / (double)segmentCount;
            var parameterLength = endParameter - startParameter;
            var start = BezierValue(nativePoints, startParameter);
            var end = BezierValue(nativePoints, endParameter);
            var startDerivative = BezierDerivative(nativePoints, startParameter);
            var endDerivative = BezierDerivative(nativePoints, endParameter);
            var expectedControl1X = start.X + parameterLength * startDerivative.X / 3d;
            var expectedControl1Y = start.Y + parameterLength * startDerivative.Y / 3d;
            var expectedControl2X = end.X - parameterLength * endDerivative.X / 3d;
            var expectedControl2Y = end.Y - parameterLength * endDerivative.Y / 3d;
            var actual = commands[index + 1].CubicBezierTo;
            // Authored lowering rounds each coordinate to 1/1000 point native
            // units. A small tolerance accounts for that quantization while
            // rejecting an arbitrary multi-segment cubic path.
            if (!CloseEnough(actual.Control1.X, expectedControl1X) ||
                !CloseEnough(actual.Control1.Y, expectedControl1Y) ||
                !CloseEnough(actual.Control2.X, expectedControl2X) ||
                !CloseEnough(actual.Control2.Y, expectedControl2Y))
            {
                points = [];
                return false;
            }
        }
        return true;
    }

    private static bool TryRecoverBezierControls(
        IReadOnlyList<PresentationCustomGeometryPoint> anchors,
        IReadOnlyList<PresentationCustomGeometryCommand> commands,
        out PresentationCustomGeometryPoint[] points)
    {
        points = [];
        var count = anchors.Count;
        if (count < 3 || count > MaxSmoothProjectionPoints) return false;
        if (!SolveBezierControls(anchors, commands, axis: 0, out var x) ||
            !SolveBezierControls(anchors, commands, axis: 1, out var y))
            return false;

        points = new PresentationCustomGeometryPoint[count];
        for (var index = 0; index < count; index++)
        {
            if (!double.IsFinite(x[index]) || !double.IsFinite(y[index]) ||
                x[index] < long.MinValue || x[index] > long.MaxValue ||
                y[index] < long.MinValue || y[index] > long.MaxValue)
            {
                points = [];
                return false;
            }
            points[index] = new PresentationCustomGeometryPoint
            {
                X = checked((long)Math.Round(x[index], MidpointRounding.AwayFromZero)),
                Y = checked((long)Math.Round(y[index], MidpointRounding.AwayFromZero)),
            };
        }
        // The native path quantizes every authored coordinate independently.
        // For the small point sets where compact Kimi syntax remains useful,
        // choose the nearby integer control lattice point with the smallest
        // complete (endpoint + both-control) residual. Larger sets stay
        // typed unless the rounded interpolation already proves itself below.
        if (count <= 8)
        {
            SnapControls(points, x, y, commands, axis: 0);
            SnapControls(points, x, y, commands, axis: 1);
        }
        points[0] = anchors[0];
        points[^1] = anchors[^1];

        // Reconstructing endpoint samples from the quantized controls guards
        // against accepting a numerically unstable interpolation as compact
        // Kimi syntax.
        var native = points.Select(point => (X: (double)point.X, Y: (double)point.Y)).ToArray();
        var segmentCount = count - 1;
        for (var index = 0; index < count; index++)
        {
            var expected = BezierValue(native, index / (double)segmentCount);
            if (!CloseEnough(anchors[index].X, expected.X) || !CloseEnough(anchors[index].Y, expected.Y))
            {
                points = [];
                return false;
            }
        }
        return true;
    }

    private static void SnapControls(
        PresentationCustomGeometryPoint[] points,
        IReadOnlyList<double> xEstimates,
        IReadOnlyList<double> yEstimates,
        IReadOnlyList<PresentationCustomGeometryCommand> commands,
        int axis)
    {
        var count = points.Length;
        var candidate = points.Select(point => axis == 0 ? point.X : point.Y).ToArray();
        var best = candidate.ToArray();
        var bestScore = ObservationResidual(points, commands, axis);
        var estimates = axis == 0 ? xEstimates : yEstimates;
        Search(1);
        for (var index = 0; index < count; index++)
            if (axis == 0) points[index].X = best[index];
            else points[index].Y = best[index];

        void Search(int index)
        {
            if (index >= count - 1)
            {
                var score = ObservationResidual(points, commands, axis, candidate);
                if (score < bestScore)
                {
                    bestScore = score;
                    Array.Copy(candidate, best, count);
                }
                return;
            }
            var center = checked((long)Math.Round(estimates[index], MidpointRounding.AwayFromZero));
            for (var delta = -3; delta <= 3; delta++)
            {
                candidate[index] = checked(center + delta);
                Search(index + 1);
            }
            candidate[index] = best[index];
        }
    }

    private static double ObservationResidual(
        IReadOnlyList<PresentationCustomGeometryPoint> points,
        IReadOnlyList<PresentationCustomGeometryCommand> commands,
        int axis,
        IReadOnlyList<long>? axisOverride = null)
    {
        var native = points.Select((point, index) => (
            X: (double)(axis == 0 && axisOverride is not null ? axisOverride[index] : point.X),
            Y: (double)(axis == 1 && axisOverride is not null ? axisOverride[index] : point.Y))).ToArray();
        var segmentCount = native.Length - 1;
        var score = 0d;
        for (var index = 0; index < native.Length; index++)
        {
            var expected = BezierValue(native, index / (double)segmentCount);
            var actual = index == 0
                ? commands[0].MoveTo
                : commands[index].CubicBezierTo.End;
            var difference = (axis == 0 ? actual.X : actual.Y) - (axis == 0 ? expected.X : expected.Y);
            score += difference * difference;
        }
        for (var index = 0; index < segmentCount; index++)
        {
            var startParameter = index / (double)segmentCount;
            var endParameter = (index + 1) / (double)segmentCount;
            var parameterLength = endParameter - startParameter;
            var start = BezierValue(native, startParameter);
            var end = BezierValue(native, endParameter);
            var startDerivative = BezierDerivative(native, startParameter);
            var endDerivative = BezierDerivative(native, endParameter);
            var actual = commands[index + 1].CubicBezierTo;
            var expectedControl1 = axis == 0
                ? start.X + parameterLength * startDerivative.X / 3d
                : start.Y + parameterLength * startDerivative.Y / 3d;
            var expectedControl2 = axis == 0
                ? end.X - parameterLength * endDerivative.X / 3d
                : end.Y - parameterLength * endDerivative.Y / 3d;
            var actualControl1 = axis == 0 ? actual.Control1.X : actual.Control1.Y;
            var actualControl2 = axis == 0 ? actual.Control2.X : actual.Control2.Y;
            score += (actualControl1 - expectedControl1) * (actualControl1 - expectedControl1);
            score += (actualControl2 - expectedControl2) * (actualControl2 - expectedControl2);
        }
        return score;
    }

    private static bool SolveBezierControls(
        IReadOnlyList<PresentationCustomGeometryPoint> anchors,
        IReadOnlyList<PresentationCustomGeometryCommand> commands,
        int axis,
        out double[] controls)
    {
        var count = anchors.Count;
        controls = [];
        if (count < 3 || commands.Count != count) return false;
        var degree = count - 1;
        var rows = new List<double[]>();
        var samples = new List<double>();
        for (var row = 0; row < count; row++)
        {
            var parameter = row / (double)degree;
            rows.Add(BernsteinBasis(degree, parameter));
            samples.Add(axis == 0 ? anchors[row].X : anchors[row].Y);
        }
        for (var index = 0; index < degree; index++)
        {
            var startParameter = index / (double)degree;
            var endParameter = (index + 1) / (double)degree;
            var parameterLength = endParameter - startParameter;
            var endBasis = BernsteinBasis(degree, endParameter);
            var endDerivative = BernsteinDerivativeBasis(degree, endParameter);
            // The endpoint and second-control observations form a stable
            // bounded interpolation basis. The first-control observations are
            // still checked below, but are not included in the reconstruction
            // solve because their quantization amplifies high-degree noise.
            rows.Add(endBasis.Zip(endDerivative, (value, derivative) => value - parameterLength * derivative / 3d).ToArray());
            samples.Add(axis == 0 ? commands[index + 1].CubicBezierTo.Control2.X : commands[index + 1].CubicBezierTo.Control2.Y);
        }

        var normal = new double[count, count + 1];
        for (var row = 0; row < rows.Count; row++)
            for (var column = 0; column < count; column++)
            {
                var coefficient = rows[row][column];
                if (!double.IsFinite(coefficient)) return false;
                for (var target = 0; target < count; target++) normal[column, target] += coefficient * rows[row][target];
                normal[column, count] += coefficient * samples[row];
            }
        if (!SolveLinearSystem(normal, count, out controls)) return false;
        return controls.All(double.IsFinite);
    }

    private static double[] BernsteinBasis(int degree, double parameter)
    {
        var basis = new double[degree + 1];
        if (parameter <= 0)
        {
            basis[0] = 1;
            return basis;
        }
        if (parameter >= 1)
        {
            basis[degree] = 1;
            return basis;
        }
        basis[0] = Math.Pow(1 - parameter, degree);
        for (var index = 1; index <= degree; index++)
            basis[index] = basis[index - 1] * (degree - index + 1d) / index * parameter / (1 - parameter);
        return basis;
    }

    private static double[] BernsteinDerivativeBasis(int degree, double parameter)
    {
        var basis = new double[degree + 1];
        var lower = BernsteinBasis(degree - 1, parameter);
        for (var index = 0; index <= degree; index++)
            basis[index] = degree * ((index == 0 ? 0 : lower[index - 1]) - (index == degree ? 0 : lower[index]));
        return basis;
    }

    private static bool SolveLinearSystem(double[,] matrix, int count, out double[] result)
    {
        result = [];
        for (var column = 0; column < count; column++)
        {
            var pivot = column;
            for (var row = column + 1; row < count; row++)
                if (Math.Abs(matrix[row, column]) > Math.Abs(matrix[pivot, column])) pivot = row;
            if (!double.IsFinite(matrix[pivot, column]) || Math.Abs(matrix[pivot, column]) < 1e-12)
                return false;
            if (pivot != column)
                for (var index = column; index <= count; index++)
                    (matrix[column, index], matrix[pivot, index]) = (matrix[pivot, index], matrix[column, index]);
            var divisor = matrix[column, column];
            for (var index = column; index <= count; index++) matrix[column, index] /= divisor;
            for (var row = 0; row < count; row++)
            {
                if (row == column) continue;
                var factor = matrix[row, column];
                if (Math.Abs(factor) < 1e-15) continue;
                for (var index = column; index <= count; index++) matrix[row, index] -= factor * matrix[column, index];
            }
        }
        result = Enumerable.Range(0, count).Select(index => matrix[index, count]).ToArray();
        return result.All(double.IsFinite);
    }

    private static bool CloseEnough(long actual, double expected) => Math.Abs(actual - expected) <= 3;

    private static bool IsLiteral(PresentationCustomGeometryPoint point) =>
        !point.HasXReference && !point.HasYReference;

    internal static void Apply(PresentationShape target, JsonElement source, string elementId)
    {
        if (source.ValueKind != JsonValueKind.Object ||
            !source.TryGetProperty("viewBox", out var viewBox) ||
            !source.TryGetProperty("commands", out var commands) ||
            viewBox.ValueKind != JsonValueKind.Object || commands.ValueKind != JsonValueKind.Array)
            throw new CodecException("ppj.line.path", $"Line {elementId} has an invalid path envelope.");

        var originX = Number(viewBox, "x", elementId);
        var originY = Number(viewBox, "y", elementId);
        var width = Number(viewBox, "width", elementId);
        var height = Number(viewBox, "height", elementId);
        if (width <= 0 || height <= 0 || width > 100_000 || height > 100_000)
            throw new CodecException("ppj.line.path", $"Line {elementId} path viewBox must have positive bounded dimensions.");

        var path = new PresentationCustomGeometryPath
        {
            Width = Coordinate(width),
            Height = Coordinate(height),
            FillMode = PresentationCustomGeometryPath.Types.FillMode.None,
            Stroke = true,
        };
        var hasCurrentPoint = false;
        var hasSubpathStart = false;
        var count = 0;
        foreach (var command in commands.EnumerateArray())
        {
            if (++count > 10_000)
                throw new CodecException("presentation_item_budget_exceeded", $"Line {elementId} exceeds the path command budget.");
            var op = command.TryGetProperty("op", out var operation) ? operation.GetString() : null;
            var native = new PresentationCustomGeometryCommand();
            switch (op)
            {
                case "moveTo":
                    native.MoveTo = Point(command, originX, originY, elementId);
                    hasCurrentPoint = true;
                    hasSubpathStart = true;
                    break;
                case "lineTo":
                    RequireCurrent(hasCurrentPoint, elementId, op);
                    native.LineTo = Point(command, originX, originY, elementId);
                    hasCurrentPoint = true;
                    break;
                case "quadraticTo":
                    RequireCurrent(hasCurrentPoint, elementId, op);
                    native.QuadraticBezierTo = new PresentationCustomGeometryQuadraticBezier
                    {
                        Control = Point(command, originX, originY, elementId, "x1", "y1"),
                        End = Point(command, originX, originY, elementId),
                    };
                    hasCurrentPoint = true;
                    break;
                case "cubicTo":
                    RequireCurrent(hasCurrentPoint, elementId, op);
                    native.CubicBezierTo = new PresentationCustomGeometryCubicBezier
                    {
                        Control1 = Point(command, originX, originY, elementId, "x1", "y1"),
                        Control2 = Point(command, originX, originY, elementId, "x2", "y2"),
                        End = Point(command, originX, originY, elementId),
                    };
                    hasCurrentPoint = true;
                    break;
                case "arcTo":
                    RequireCurrent(hasCurrentPoint, elementId, op);
                    native.ArcTo = new PresentationCustomGeometryArc
                    {
                        WidthRadius = Coordinate(Number(command, "radiusX", elementId)),
                        HeightRadius = Coordinate(Number(command, "radiusY", elementId)),
                        StartAngle = Angle(Number(command, "startAngle", elementId)),
                        SweepAngle = Angle(Number(command, "sweepAngle", elementId)),
                    };
                    hasCurrentPoint = true;
                    break;
                case "close":
                    throw new CodecException("ppj.line.path", $"Line {elementId} paths cannot close a stroked subpath.");
                default:
                    throw new CodecException("ppj.line.path", $"Line {elementId} contains unsupported path operation {op ?? "(missing)"}.");
            }
            path.Commands.Add(native);
        }
        if (path.Commands.Count < 2 || path.Commands[0].CommandCase != PresentationCustomGeometryCommand.CommandOneofCase.MoveTo ||
            !hasSubpathStart)
            throw new CodecException("ppj.line.path", $"Line {elementId} requires a moveTo followed by at least one drawing command.");

        target.Geometry = "custom";
        target.CustomPaths.Clear();
        target.CustomAdjustments.Clear();
        target.CustomGuides.Clear();
        target.CustomConnectionSites.Clear();
        target.CustomAdjustmentHandles.Clear();
        target.TextRectangle = null;
        target.CustomPaths.Add(path);
    }

    private static JsonObject ProjectCommand(PresentationCustomGeometryCommand command) => command.CommandCase switch
    {
        PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => Point("moveTo", command.MoveTo),
        PresentationCustomGeometryCommand.CommandOneofCase.LineTo => Point("lineTo", command.LineTo),
        PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo => new JsonObject
        {
            ["op"] = "quadraticTo",
            ["x1"] = PointValue(command.QuadraticBezierTo.Control.X),
            ["y1"] = PointValue(command.QuadraticBezierTo.Control.Y),
            ["x"] = PointValue(command.QuadraticBezierTo.End.X),
            ["y"] = PointValue(command.QuadraticBezierTo.End.Y),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo => new JsonObject
        {
            ["op"] = "cubicTo",
            ["x1"] = PointValue(command.CubicBezierTo.Control1.X),
            ["y1"] = PointValue(command.CubicBezierTo.Control1.Y),
            ["x2"] = PointValue(command.CubicBezierTo.Control2.X),
            ["y2"] = PointValue(command.CubicBezierTo.Control2.Y),
            ["x"] = PointValue(command.CubicBezierTo.End.X),
            ["y"] = PointValue(command.CubicBezierTo.End.Y),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => new JsonObject
        {
            ["op"] = "arcTo",
            ["radiusX"] = PointValue(command.ArcTo.WidthRadius),
            ["radiusY"] = PointValue(command.ArcTo.HeightRadius),
            ["startAngle"] = command.ArcTo.StartAngle / 60_000d,
            ["sweepAngle"] = command.ArcTo.SweepAngle / 60_000d,
        },
        _ => throw new CodecException("unsupported_presentation_line", "The native line path contains an unsupported command."),
    };

    private static JsonObject Point(string op, PresentationCustomGeometryPoint point) => new()
    {
        ["op"] = op,
        ["x"] = PointValue(point.X),
        ["y"] = PointValue(point.Y),
    };

    private static string PointText(PresentationCustomGeometryPoint point) =>
        $"{PointValue(point.X).ToString(CultureInfo.InvariantCulture)},{PointValue(point.Y).ToString(CultureInfo.InvariantCulture)}";

    private static PresentationCustomGeometryPoint Point(
        JsonElement command,
        double originX,
        double originY,
        string elementId,
        string xName = "x",
        string yName = "y") => new()
    {
        X = Coordinate(Number(command, xName, elementId) - originX),
        Y = Coordinate(Number(command, yName, elementId) - originY),
    };

    private static double Number(JsonElement owner, string name, string elementId)
    {
        if (!owner.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.Number ||
            !value.TryGetDouble(out var number) || double.IsNaN(number) || double.IsInfinity(number))
            throw new CodecException("ppj.line.path", $"Line {elementId} path field {name} must be a finite number.");
        return number;
    }

    private static double Finite(JsonElement value, string field, string elementId)
    {
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetDouble(out var number) || !double.IsFinite(number))
            throw new CodecException("ppj.line.points", $"Line {elementId} {field} must be a finite number.");
        return number;
    }

    private static bool TryFinite(string value, out double number) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) && double.IsFinite(number);

    private static JsonObject PointCommand(string operation, (double X, double Y) point) => new()
    {
        ["op"] = operation,
        ["x"] = point.X,
        ["y"] = point.Y,
    };

    private static void RequireCurrent(bool current, string elementId, string operation)
    {
        if (!current)
            throw new CodecException("ppj.line.path", $"Line {elementId} operation {operation} requires a preceding moveTo.");
    }

    private static long Coordinate(double value) => checked((long)Math.Round(value * UnitsPerPoint, MidpointRounding.AwayFromZero));
    private static int Angle(double value) => checked((int)Math.Round(value * 60_000d, MidpointRounding.AwayFromZero));
    private static double PointValue(long value) => value / UnitsPerPoint;
}
