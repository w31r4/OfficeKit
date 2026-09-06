using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Bounded DrawingML custom paths used by source-built presentation templates.
// Coordinates and arc values may reference DrawingML built-ins or one ordered
// adjustment/guide graph;
// formula parsing and evaluation stay in PptxCustomGeometryFormulaCodec. Shape-
// local text bounds are delegated to a leaf that retains the private numeric
// scaling profile while accepting standard literal/reference edges. XY/polar
// adjustment handles and connection sites share the same evaluated graph. 3D
// and relative lighten/darken path fill remain opaque and fail closed. Connection-site and
// handle array positions are native identity; handle kind and controlled guide
// references are identity too.
internal static class PptxCustomGeometryCodec
{
    private const int MaxPaths = 64;
    private const int MaxCommands = 16_384;
    private const int MaxConnectionSites = 1_024;
    private const long MaxCoordinate = int.MaxValue;
    private const int FullTurnAngle = 21_600_000;
    private sealed class Profile
    {
        internal required A.PathList Paths { get; init; }
        internal required PptxCustomGeometryFormulaCodec.Graph Formulas { get; init; }
        internal required IReadOnlyList<PresentationCustomGeometryConnectionSite> ConnectionSites { get; init; }
        internal required IReadOnlyList<PresentationCustomGeometryAdjustmentHandle> AdjustmentHandles { get; init; }
        internal PresentationCustomGeometryTextRectangle? TextRectangle { get; init; }
    }

    internal static bool Supports(A.CustomGeometry? geometry, long widthEmu, long heightEmu)
    {
        return TryProfile(geometry, widthEmu, heightEmu, out _);
    }

    internal static bool TryReadCanonicalGuideFormula(
        A.CustomGeometry? geometry,
        long widthEmu,
        long heightEmu,
        uint nativeIndex,
        out string formula)
    {
        formula = string.Empty;
        if (!TryProfile(geometry, widthEmu, heightEmu, out var profile) ||
            nativeIndex >= (uint)profile.Formulas.Guides.Count)
            return false;
        formula = profile.Formulas.Guides[(int)nativeIndex].Formula;
        return formula.Length > 0;
    }

    internal static bool TryReadCanonicalAdjustmentFormula(
        A.CustomGeometry? geometry,
        long widthEmu,
        long heightEmu,
        uint nativeIndex,
        out string formula)
    {
        formula = string.Empty;
        if (!TryProfile(geometry, widthEmu, heightEmu, out var profile) ||
            nativeIndex >= (uint)profile.Formulas.Adjustments.Count)
            return false;
        formula = profile.Formulas.Adjustments[(int)nativeIndex].Formula;
        return formula.Length > 0;
    }

    internal static bool TryPreset(string name, out A.ShapeTypeValues preset) =>
        PptxPresetGeometryAdjustmentCodec.TryPreset(name, out preset);

    internal static bool TryPresetName(A.ShapeTypeValues preset, out string name) =>
        PptxPresetGeometryAdjustmentCodec.TryPresetName(preset, out name);

    internal static void Read(A.CustomGeometry? geometry, long widthEmu, long heightEmu, PresentationShape target)
    {
        if (!TryProfile(geometry, widthEmu, heightEmu, out var profile))
        {
            // A source-bound image-filled shape may carry a legal custom path
            // graph whose adjustment formulas are outside our executable
            // guide grammar.  Keep the graph source-owned, but retain the
            // direct adjustment list so independently literal `val N`
            // siblings can still receive native-leaf capabilities.  Do not
            // apply this fallback to picture masks or source-free shapes:
            // ImageFillAssetId is populated only for the bounded shape-level
            // image-fill profile.
            if (!string.IsNullOrWhiteSpace(target.ImageFillAssetId) &&
                TryReadSourceBoundAdjustmentGuides(geometry, out var adjustments))
                target.CustomAdjustments.Add(adjustments);
            return;
        }
        target.CustomAdjustments.Add(profile.Formulas.Adjustments);
        target.CustomGuides.Add(profile.Formulas.Guides);
        target.CustomConnectionSites.Add(profile.ConnectionSites);
        target.CustomAdjustmentHandles.Add(profile.AdjustmentHandles);
        target.TextRectangle = profile.TextRectangle;
        foreach (var nativePath in profile.Paths.Elements<A.Path>())
        {
            var path = new PresentationCustomGeometryPath
            {
                Width = checked((long)(nativePath.Width?.Value ?? 0)),
                Height = checked((long)(nativePath.Height?.Value ?? 0)),
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

    // Narrow source-bound fallback for an adjustment list that is structurally
    // simple even when the rest of custGeom (or one sibling formula) is not
    // evaluable by the closed formula codec.  The returned guides preserve
    // their native order and exact formula text; only direct `val N` entries
    // are later exposed as editable leaves.  Unknown children, missing
    // identity attributes, duplicate names and oversized lists are rejected.
    private static bool TryReadSourceBoundAdjustmentGuides(
        A.CustomGeometry? geometry,
        out IReadOnlyList<PresentationCustomGeometryGuide> adjustments)
    {
        adjustments = [];
        if (geometry is null || geometry.HasAttributes || geometry.ChildElements.Count == 0 ||
            geometry.FirstChild is not A.AdjustValueList list || list.HasAttributes)
            return false;
        var nativeGuides = list.Elements<A.ShapeGuide>().ToArray();
        if (nativeGuides.Length == 0 || nativeGuides.Length > 256 ||
            list.ChildElements.Count != nativeGuides.Length)
            return false;
        var result = new List<PresentationCustomGeometryGuide>(nativeGuides.Length);
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var native in nativeGuides)
        {
            if (native.ChildElements.Count != 0 ||
                !HasOnlyAttributes(native, "name", "fmla") ||
                native.Name?.Value is not { Length: > 0 } name ||
                native.Formula?.Value is not { Length: > 0 } formula ||
                formula.Length > 256 || !names.Add(name))
                return false;
            result.Add(new PresentationCustomGeometryGuide { Name = name, Formula = formula });
        }
        adjustments = result;
        return true;
    }

    private static bool TryProfile(A.CustomGeometry? geometry, long widthEmu, long heightEmu, out Profile profile)
    {
        profile = null!;
        if (geometry is null || geometry.HasAttributes || geometry.ChildElements.Count < 1)
            return false;
        var index = 0;
        var adjustments = geometry.ChildElements[index] is A.AdjustValueList sourceAdjustments
            ? sourceAdjustments
            : new A.AdjustValueList();
        if (geometry.ChildElements[index] is A.AdjustValueList) index++;
        A.ShapeGuideList? guideList = null;
        if (index < geometry.ChildElements.Count && geometry.ChildElements[index] is A.ShapeGuideList sourceGuides)
        {
            if (sourceGuides.HasAttributes) return false;
            guideList = sourceGuides;
            index++;
        }
        A.AdjustHandleList? nativeAdjustmentHandles = null;
        if (index < geometry.ChildElements.Count && geometry.ChildElements[index] is A.AdjustHandleList handles)
        {
            nativeAdjustmentHandles = handles;
            index++;
        }
        A.ConnectionSiteList? nativeConnectionSites = null;
        if (index < geometry.ChildElements.Count && geometry.ChildElements[index] is A.ConnectionSiteList connections)
        {
            nativeConnectionSites = connections;
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
        if (!PptxCustomGeometryTextRectangleCodec.TryPrepare(nativeRectangle, allGuides, widthEmu, heightEmu, out var textRectangleProfile, out var userGuideCount))
            return false;
        if (!PptxCustomGeometryFormulaCodec.TryRead(adjustments, allGuides.Take(userGuideCount), widthEmu, heightEmu, out var formulas))
            return false;
        if (!PptxCustomGeometryTextRectangleCodec.TryRead(textRectangleProfile, formulas, out var textRectangle))
            return false;
        if (!PptxCustomGeometryHandleCodec.TryRead(nativeAdjustmentHandles, formulas, widthEmu, heightEmu, out var adjustmentHandles))
            return false;
        if (!TryReadConnectionSites(nativeConnectionSites, formulas, widthEmu, heightEmu, out var connectionSites))
            return false;
        var paths = pathList.Elements<A.Path>().ToArray();
        if (paths.Length is < 1 or > MaxPaths || pathList.ChildElements.Count != paths.Length) return false;
        var commandCount = 0;
        if (!paths.All(path => Supports(path, formulas, ref commandCount))) return false;
        profile = new Profile { Paths = pathList, Formulas = formulas, ConnectionSites = connectionSites, AdjustmentHandles = adjustmentHandles, TextRectangle = textRectangle };
        return true;
    }

    internal static void Validate(PresentationShape shape, string shapeId)
    {
        if (shape.Geometry != "custom")
        {
            if (!PptxPresetGeometryAdjustmentCodec.HasProfile(shape.Geometry))
                throw new CodecException("unsupported_presentation_geometry", $"Presentation shape {shapeId} uses unsupported preset geometry {shape.Geometry}.");
            if (shape.CustomPaths.Count > 0 || shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 || shape.CustomConnectionSites.Count > 0 || shape.CustomAdjustmentHandles.Count > 0 || shape.TextRectangle is not null)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has custom geometry data without custom geometry.");
            PptxPresetGeometryAdjustmentCodec.Validate(shape.Geometry, shape.PresetAdjustments, shapeId);
            return;
        }
        if (shape.PresetAdjustments.Count > 0)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has preset adjustments with custom geometry.");
        // A source-bound image-filled shape may intentionally carry a custom
        // path graph outside this codec's semantic profile.  Its image fill,
        // native path tokens and (when present) adjustment formulas remain
        // source-owned; the bounded frame and independently literal
        // adjustment leaves can still be edited without pretending that the
        // opaque graph is authorable.  Source-free image fills are rejected
        // by ValidatePresentationElement before reaching this branch.
        if (shape.CustomPaths.Count == 0 &&
            !string.IsNullOrWhiteSpace(shape.ImageFillAssetId) &&
            shape.CustomGuides.Count == 0 &&
            shape.CustomConnectionSites.Count == 0 &&
            shape.CustomAdjustmentHandles.Count == 0 &&
             shape.TextRectangle is null)
             return;
        if (shape.CustomPaths.Count is < 1 or > MaxPaths)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} custom geometry must contain 1 through {MaxPaths} paths.");
        var formulas = PptxCustomGeometryFormulaCodec.Validate(shape, shapeId);
        PptxCustomGeometryHandleCodec.Validate(shape, shapeId, formulas);
        ValidateConnectionSites(shape, shapeId, formulas);
        PptxCustomGeometryTextRectangleCodec.Validate(shape.TextRectangle, shapeId, formulas);
        var commandCount = 0;
        foreach (var path in shape.CustomPaths)
        {
            if (path.Width is < 0 or > MaxCoordinate || path.Height is < 0 or > MaxCoordinate || path.Commands.Count == 0)
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

    internal static void Apply(P.ShapeProperties properties, PresentationShape shape, string shapeId)
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
            if (!PptxPresetGeometryAdjustmentCodec.TryPreset(shape.Geometry, out var presetGeometry))
                throw new CodecException("unsupported_presentation_geometry", $"Presentation shape {shapeId} uses unsupported preset geometry {shape.Geometry}.");
            preset.Preset = presetGeometry;
            PptxPresetGeometryAdjustmentCodec.Apply(preset, shape.Geometry, shape.PresetAdjustments, shapeId);
            return;
        }
        var transform = properties.GetFirstChild<A.Transform2D>();
        var widthEmu = transform?.Extents?.Cx?.Value ?? shape.WidthEmu;
        var heightEmu = transform?.Extents?.Cy?.Value ?? shape.HeightEmu;
        var existingGeometry = properties.GetFirstChild<A.CustomGeometry>();
        var omitAdjustmentList = existingGeometry is not null &&
            existingGeometry.ChildElements.Count > 0 &&
            existingGeometry.ChildElements[0] is not A.AdjustValueList &&
            shape.CustomAdjustments.Count == 0;
        if (existingGeometry is not null && TryProfile(existingGeometry, widthEmu, heightEmu, out var existingProfile))
        {
            if (existingProfile.ConnectionSites.Count != shape.CustomConnectionSites.Count)
                throw new CodecException("unsupported_presentation_edit", $"Source-preserving PPTX export requires custom shape connection-site list length to remain fixed at {existingProfile.ConnectionSites.Count}; each existing index is the native identity.");
            if (!PptxCustomGeometryHandleCodec.TopologyEquals(existingProfile.AdjustmentHandles, shape.CustomAdjustmentHandles))
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export requires the original custom-shape adjustment-handle order, kind, and controlled adjustment identity.");
        }
        properties.GetFirstChild<A.PresetGeometry>()?.Remove();
        existingGeometry?.Remove();
        OpenXmlElement geometry = Build(shape, widthEmu, heightEmu, shapeId, omitAdjustmentList);
        if (transform is null) properties.PrependChild(geometry);
        else properties.InsertAfter(geometry, transform);
    }

    private static A.CustomGeometry Build(PresentationShape shape, long widthEmu, long heightEmu, string shapeId, bool omitAdjustmentList = false)
    {
        var paths = new A.PathList();
        foreach (var source in shape.CustomPaths)
        {
            var path = new A.Path();
            if (source.Width > 0) path.Width = source.Width;
            if (source.Height > 0) path.Height = source.Height;
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
        var geometry = new A.CustomGeometry();
        if (!omitAdjustmentList) geometry.Append(adjustments);
        var formulaGraph = PptxCustomGeometryFormulaCodec.Validate(shape, shapeId);
        var textRectangle = shape.TextRectangle is null
            ? null
            : PptxCustomGeometryTextRectangleCodec.Build(shape.TextRectangle, widthEmu, heightEmu, formulaGraph, shapeId);
        A.ShapeGuideList? guides = null;
        if (shape.CustomGuides.Count > 0 || textRectangle?.Guides.Count > 0)
        {
            guides = new A.ShapeGuideList(shape.CustomGuides.Select(PptxCustomGeometryFormulaCodec.Write));
            if (textRectangle is not null) guides.Append(textRectangle.Guides);
            geometry.Append(guides);
        }
        if (shape.CustomAdjustmentHandles.Count > 0)
            geometry.Append(PptxCustomGeometryHandleCodec.Build(shape.CustomAdjustmentHandles));
        if (shape.CustomConnectionSites.Count > 0)
            geometry.Append(new A.ConnectionSiteList(shape.CustomConnectionSites.Select(ConnectionSite)));
        if (textRectangle is not null) geometry.Append(textRectangle.Rectangle);
        geometry.Append(paths);
        return geometry;
    }

    private static bool Supports(A.Path path, PptxCustomGeometryFormulaCodec.Graph formulas, ref int commandCount)
    {
        if (path.Width is { HasValue: false } || path.Width?.Value > MaxCoordinate ||
            path.Height is { HasValue: false } || path.Height?.Value > MaxCoordinate ||
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

    private static bool TryReadConnectionSites(
        A.ConnectionSiteList? source,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        long widthEmu,
        long heightEmu,
        out IReadOnlyList<PresentationCustomGeometryConnectionSite> connectionSites)
    {
        connectionSites = [];
        if (source is null) return true;
        var nativeSites = source.Elements<A.ConnectionSite>().ToArray();
        if (source.HasAttributes || source.ChildElements.Count != nativeSites.Length || nativeSites.Length > MaxConnectionSites)
            return false;
        var result = new List<PresentationCustomGeometryConnectionSite>(nativeSites.Length);
        foreach (var nativeSite in nativeSites)
        {
            if (!HasOnlyAttributes(nativeSite, "ang") || nativeSite.ChildElements.Count != 1 ||
                nativeSite.ChildElements[0] is not A.Position position ||
                !HasOnlyAttributes(position, "x", "y") || !HasNoInnerXml(position) ||
                !TryValue(nativeSite.Angle?.Value, formulas, out var angle) || Math.Abs(angle) > FullTurnAngle ||
                !TryValue(position.X?.Value, formulas, out var x) || x < 0 || x > widthEmu ||
                !TryValue(position.Y?.Value, formulas, out var y) || y < 0 || y > heightEmu)
                return false;
            var site = new PresentationCustomGeometryConnectionSite();
            if (TryAngle(nativeSite.Angle?.Value, out var literalAngle)) site.Angle60000 = literalAngle;
            else site.AngleReference = nativeSite.Angle!.Value!;
            if (TryCoordinate(position.X?.Value, out var literalX)) site.XEmu = literalX;
            else site.XReference = position.X!.Value!;
            if (TryCoordinate(position.Y?.Value, out var literalY)) site.YEmu = literalY;
            else site.YReference = position.Y!.Value!;
            result.Add(site);
        }
        connectionSites = result;
        return true;
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

    private static A.ConnectionSite ConnectionSite(PresentationCustomGeometryConnectionSite source) => new(
        new A.Position
        {
            X = source.HasXReference ? source.XReference : source.XEmu.ToString(CultureInfo.InvariantCulture),
            Y = source.HasYReference ? source.YReference : source.YEmu.ToString(CultureInfo.InvariantCulture),
        })
    {
        Angle = source.HasAngleReference ? source.AngleReference : source.Angle60000.ToString(CultureInfo.InvariantCulture),
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

    private static void ValidateConnectionSites(
        PresentationShape shape,
        string shapeId,
        PptxCustomGeometryFormulaCodec.Graph formulas)
    {
        if (shape.CustomConnectionSites.Count > MaxConnectionSites)
            throw new CodecException("presentation_item_budget_exceeded", $"Presentation shape {shapeId} custom geometry exceeds the {MaxConnectionSites}-connection-site budget.");
        foreach (var site in shape.CustomConnectionSites)
        {
            if (!TryWireValue(site.HasAngleReference, site.AngleReference, site.Angle60000, formulas, out var angle) ||
                Math.Abs(angle) > FullTurnAngle ||
                !TryWireValue(site.HasXReference, site.XReference, site.XEmu, formulas, out var x) || x < 0 || x > shape.WidthEmu ||
                !TryWireValue(site.HasYReference, site.YReference, site.YEmu, formulas, out var y) || y < 0 || y > shape.HeightEmu)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid bounded custom connection site or formula reference.");
        }
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
