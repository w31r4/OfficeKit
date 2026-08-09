using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns the a:rect coordinate protocol and OfficeKit's private numeric-scaling
// guide tail. The parent codec only decides custGeom child order; this leaf
// recognizes native literal/reference edges, hides the private guide profile,
// validates resolved bounds, and emits one canonical representation.
internal static class PptxCustomGeometryTextRectangleCodec
{
    private const long MaxCoordinate = int.MaxValue;

    private readonly record struct Edge(
        string Attribute,
        string Guide,
        string Axis,
        Func<PresentationCustomGeometryTextRectangle, long> Literal,
        Func<PresentationCustomGeometryTextRectangle, bool> HasReference,
        Func<PresentationCustomGeometryTextRectangle, string> Reference,
        Action<PresentationCustomGeometryTextRectangle, long> SetLiteral,
        Action<PresentationCustomGeometryTextRectangle, string> SetReference);

    private static readonly Edge[] Edges =
    [
        new("l", "officeKitTextLeft", "w", item => item.LeftEmu, item => item.HasLeftReference, item => item.LeftReference, (item, value) => item.LeftEmu = value, (item, value) => item.LeftReference = value),
        new("t", "officeKitTextTop", "h", item => item.TopEmu, item => item.HasTopReference, item => item.TopReference, (item, value) => item.TopEmu = value, (item, value) => item.TopReference = value),
        new("r", "officeKitTextRight", "w", item => item.RightEmu, item => item.HasRightReference, item => item.RightReference, (item, value) => item.RightEmu = value, (item, value) => item.RightReference = value),
        new("b", "officeKitTextBottom", "h", item => item.BottomEmu, item => item.HasBottomReference, item => item.BottomReference, (item, value) => item.BottomEmu = value, (item, value) => item.BottomReference = value),
    ];

    internal sealed class NativeProfile
    {
        internal required IReadOnlyList<string> Values { get; init; }
        internal IReadOnlyList<long>? PrivateCoordinates { get; init; }
    }

    internal sealed record Output(IReadOnlyList<A.ShapeGuide> Guides, A.Rectangle Rectangle);

    internal static bool TryPrepare(
        A.Rectangle? source,
        IReadOnlyList<A.ShapeGuide> guides,
        long widthEmu,
        long heightEmu,
        out NativeProfile? profile,
        out int userGuideCount)
    {
        profile = null;
        userGuideCount = guides.Count;
        if (source is null) return true;
        if (!HasNoInnerXml(source) || !HasOnlyAttributes(source, Edges.Select(edge => edge.Attribute).ToArray())) return false;

        var values = new[] { source.Left?.Value, source.Top?.Value, source.Right?.Value, source.Bottom?.Value };
        if (values.Any(string.IsNullOrEmpty)) return false;
        var usesPrivateGuide = Edges.Select((edge, index) => values[index] == edge.Guide).Any(value => value);
        IReadOnlyList<long>? privateCoordinates = null;
        if (usesPrivateGuide)
        {
            if (guides.Count < Edges.Length) return false;
            userGuideCount = guides.Count - Edges.Length;
            var decoded = new long[Edges.Length];
            for (var index = 0; index < Edges.Length; index++)
            {
                var edge = Edges[index];
                var extent = edge.Axis == "w" ? widthEmu : heightEmu;
                if (!TryScaledGuide(guides[userGuideCount + index], edge.Guide, edge.Axis, extent, out decoded[index])) return false;
            }
            privateCoordinates = decoded;
        }
        profile = new NativeProfile { Values = values!, PrivateCoordinates = privateCoordinates };
        return true;
    }

    internal static bool TryRead(
        NativeProfile? profile,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        out PresentationCustomGeometryTextRectangle? rectangle)
    {
        rectangle = null;
        if (profile is null) return true;
        var result = new PresentationCustomGeometryTextRectangle();
        var resolved = new double[Edges.Length];
        for (var index = 0; index < Edges.Length; index++)
        {
            var edge = Edges[index];
            var value = profile.Values[index];
            if (value == edge.Guide)
            {
                if (profile.PrivateCoordinates is null) return false;
                var coordinate = profile.PrivateCoordinates[index];
                edge.SetLiteral(result, coordinate);
                resolved[index] = coordinate;
            }
            else if (TryCoordinate(value, out var coordinate))
            {
                edge.SetLiteral(result, coordinate);
                resolved[index] = coordinate;
            }
            else if (formulas.TryResolveReference(value, out var evaluated))
            {
                edge.SetReference(result, value);
                resolved[index] = evaluated;
            }
            else
            {
                return false;
            }
        }
        if (resolved[0] >= resolved[2] || resolved[1] >= resolved[3]) return false;
        rectangle = result;
        return true;
    }

    internal static void Validate(
        PresentationCustomGeometryTextRectangle? rectangle,
        string shapeId,
        PptxCustomGeometryFormulaCodec.Graph formulas)
    {
        if (rectangle is null) return;
        var resolved = new double[Edges.Length];
        for (var index = 0; index < Edges.Length; index++)
        {
            var edge = Edges[index];
            var hasReference = edge.HasReference(rectangle);
            var literal = edge.Literal(rectangle);
            if (hasReference && literal != 0 ||
                !(hasReference
                    ? formulas.TryResolveReference(edge.Reference(rectangle), out resolved[index])
                    : TryLiteral(literal, out resolved[index])))
                throw Invalid(shapeId);
        }
        if (resolved[0] >= resolved[2] || resolved[1] >= resolved[3]) throw Invalid(shapeId);
    }

    internal static Output Build(
        PresentationCustomGeometryTextRectangle source,
        long widthEmu,
        long heightEmu,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        string shapeId)
    {
        Validate(source, shapeId, formulas);
        var needsPrivateGuides = Edges.Any(edge => !edge.HasReference(source));
        var guides = needsPrivateGuides
            ? Edges.Select(edge => ScaledGuide(
                edge.Guide,
                edge.HasReference(source) ? 0 : edge.Literal(source),
                edge.Axis,
                edge.Axis == "w" ? widthEmu : heightEmu)).ToArray()
            : [];
        var values = Edges.Select(edge => edge.HasReference(source) ? edge.Reference(source) : edge.Guide).ToArray();
        return new Output(guides, new A.Rectangle
        {
            Left = values[0],
            Top = values[1],
            Right = values[2],
            Bottom = values[3],
        });
    }

    private static CodecException Invalid(string shapeId) =>
        new("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid custom-geometry text rectangle or guide reference.");

    private static bool TryLiteral(long value, out double resolved)
    {
        resolved = value;
        return value is >= -MaxCoordinate and <= MaxCoordinate;
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

    private static A.ShapeGuide ScaledGuide(string name, long coordinate, string axis, long extentEmu) => new()
    {
        Name = name,
        Formula = $"*/ {coordinate.ToString(CultureInfo.InvariantCulture)} {axis} {extentEmu.ToString(CultureInfo.InvariantCulture)}",
    };

    private static bool TryCoordinate(string? value, out long coordinate) =>
        long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out coordinate) &&
        coordinate is >= -MaxCoordinate and <= MaxCoordinate;

    // OpenXmlLeafElement.HasChildren is always false even when malformed source
    // XML is retained in its shadow element. InnerXml is the lexical gate.
    private static bool HasNoInnerXml(OpenXmlElement element) => string.IsNullOrEmpty(element.InnerXml);

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }
}
