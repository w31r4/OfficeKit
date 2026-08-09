using System.Globalization;
using System.Text.RegularExpressions;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// One ordered, bounded implementation of the ECMA-376 geometry-guide grammar.
// The surrounding custom-geometry codec owns XML topology; this module owns
// names, references, operators, evaluation, and the reserved-name boundary.
internal static partial class PptxCustomGeometryFormulaCodec
{
    private const int MaxCoordinate = int.MaxValue;
    private const int MaxAdjustments = 256;
    private const int MaxGuides = 1_024;
    private const int MaxFormulaLength = 256;
    private const double AngleUnitsPerDegree = 60_000d;

    private static readonly IReadOnlyDictionary<string, int> FormulaArity = new Dictionary<string, int>(StringComparer.Ordinal)
    {
        ["*/"] = 3,
        ["+-"] = 3,
        ["+/"] = 3,
        ["?:"] = 3,
        ["abs"] = 1,
        ["at2"] = 2,
        ["cat2"] = 3,
        ["cos"] = 2,
        ["max"] = 2,
        ["min"] = 2,
        ["mod"] = 3,
        ["pin"] = 3,
        ["sat2"] = 3,
        ["sin"] = 2,
        ["sqrt"] = 1,
        ["tan"] = 2,
        ["val"] = 1,
    };

    internal sealed class Graph(
        IReadOnlyList<PresentationCustomGeometryGuide> adjustments,
        IReadOnlyList<PresentationCustomGeometryGuide> guides,
        IReadOnlyDictionary<string, double> values)
    {
        internal IReadOnlyList<PresentationCustomGeometryGuide> Adjustments { get; } = adjustments;
        internal IReadOnlyList<PresentationCustomGeometryGuide> Guides { get; } = guides;
        private IReadOnlyDictionary<string, double> Values { get; } = values;

        internal bool TryResolve(string? reference, long literal, out double value)
        {
            if (reference is not null) return Values.TryGetValue(reference, out value);
            value = literal;
            return literal is >= -MaxCoordinate and <= MaxCoordinate;
        }

        internal bool TryResolveReference(string? reference, out double value)
        {
            if (reference is not null) return Values.TryGetValue(reference, out value);
            value = 0;
            return false;
        }
    }

    private sealed class FormulaException(string message) : Exception(message);
    private readonly record struct GuideInput(string? Name, string? Formula);

    internal static bool TryRead(
        A.AdjustValueList adjustments,
        IEnumerable<A.ShapeGuide> guides,
        long widthEmu,
        long heightEmu,
        out Graph graph)
    {
        graph = null!;
        if (adjustments.HasAttributes || adjustments.ChildElements.Count != adjustments.Elements<A.ShapeGuide>().Count()) return false;
        var nativeGuides = guides.ToArray();
        if (adjustments.Elements<A.ShapeGuide>().Any(HasUnsupportedShapeGuide) || nativeGuides.Any(HasUnsupportedShapeGuide)) return false;
        try
        {
            graph = Build(
                adjustments.Elements<A.ShapeGuide>().Select(item => new GuideInput(item.Name?.Value, item.Formula?.Value)),
                nativeGuides.Select(item => new GuideInput(item.Name?.Value, item.Formula?.Value)),
                widthEmu,
                heightEmu);
            return true;
        }
        catch (FormulaException)
        {
            return false;
        }
    }

    internal static Graph Validate(PresentationShape shape, string shapeId)
    {
        try
        {
            return Build(
                shape.CustomAdjustments.Select(item => new GuideInput(item.Name, item.Formula)),
                shape.CustomGuides.Select(item => new GuideInput(item.Name, item.Formula)),
                shape.WidthEmu,
                shape.HeightEmu);
        }
        catch (FormulaException error)
        {
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid custom geometry formula graph: {error.Message}");
        }
    }

    internal static A.ShapeGuide Write(PresentationCustomGeometryGuide source) => new()
    {
        Name = source.Name,
        Formula = source.Formula,
    };

    private static Graph Build(
        IEnumerable<GuideInput> adjustmentSource,
        IEnumerable<GuideInput> guideSource,
        long widthEmu,
        long heightEmu)
    {
        var adjustments = adjustmentSource.ToArray();
        var guides = guideSource.ToArray();
        if (adjustments.Length > MaxAdjustments) throw new FormulaException($"adjustments exceed the {MaxAdjustments}-guide budget");
        if (guides.Length > MaxGuides) throw new FormulaException($"guides exceed the {MaxGuides}-guide budget");
        var values = Builtins(widthEmu, heightEmu);
        var normalizedAdjustments = Normalize(adjustments, "adjustment", values);
        var normalizedGuides = Normalize(guides, "guide", values);
        return new Graph(normalizedAdjustments, normalizedGuides, values);
    }

    private static IReadOnlyList<PresentationCustomGeometryGuide> Normalize(
        IReadOnlyList<GuideInput> source,
        string kind,
        Dictionary<string, double> values)
    {
        var result = new List<PresentationCustomGeometryGuide>(source.Count);
        for (var index = 0; index < source.Count; index++)
        {
            var name = source[index].Name ?? string.Empty;
            if (!GuideNameRegex().IsMatch(name) || IntegerRegex().IsMatch(name))
                throw new FormulaException($"{kind} {index + 1} has an invalid name");
            if (name.StartsWith("officeKit", StringComparison.OrdinalIgnoreCase))
                throw new FormulaException($"{kind} {name} uses the reserved officeKit prefix");
            if (values.ContainsKey(name)) throw new FormulaException($"{kind} {name} duplicates a built-in or earlier guide");
            var formula = NormalizeFormula(source[index].Formula, values, $"{kind} {name}");
            var value = Evaluate(formula, values, $"{kind} {name}");
            result.Add(new PresentationCustomGeometryGuide { Name = name, Formula = formula });
            values.Add(name, value);
        }
        return result;
    }

    private static string NormalizeFormula(string? source, IReadOnlyDictionary<string, double> values, string label)
    {
        if (source is null) throw new FormulaException($"{label} has no formula");
        var formula = string.Join(' ', source.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        if (formula.Length is < 1 or > MaxFormulaLength) throw new FormulaException($"{label} formula exceeds the {MaxFormulaLength}-character budget");
        var tokens = formula.Split(' ');
        if (!FormulaArity.TryGetValue(tokens[0], out var arity)) throw new FormulaException($"{label} uses unsupported operator {tokens[0]}");
        if (tokens.Length != arity + 1) throw new FormulaException($"{label} operator {tokens[0]} requires {arity} operands");
        foreach (var token in tokens.Skip(1))
        {
            if (IntegerRegex().IsMatch(token))
            {
                if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var literal) || literal is < -MaxCoordinate or > MaxCoordinate)
                    throw new FormulaException($"{label} contains an out-of-range literal");
                continue;
            }
            if (!values.ContainsKey(token)) throw new FormulaException($"{label} references unknown or forward guide {token}");
        }
        return formula;
    }

    private static double Evaluate(string formula, IReadOnlyDictionary<string, double> values, string label)
    {
        var tokens = formula.Split(' ');
        var args = tokens.Skip(1).Select(token => Operand(token, values)).ToArray();
        var result = tokens[0] switch
        {
            "*/" => args[2] == 0 ? throw new FormulaException($"{label} divides by zero") : args[0] * args[1] / args[2],
            "+-" => args[0] + args[1] - args[2],
            "+/" => args[2] == 0 ? throw new FormulaException($"{label} divides by zero") : (args[0] + args[1]) / args[2],
            "?:" => args[0] > 0 ? args[1] : args[2],
            "abs" => Math.Abs(args[0]),
            "at2" => Math.Atan2(args[1], args[0]) * 180d / Math.PI * AngleUnitsPerDegree,
            "cat2" => args[0] * Math.Cos(Math.Atan2(args[2], args[1])),
            "cos" => args[0] * Math.Cos(AngleRadians(args[1])),
            "max" => Math.Max(args[0], args[1]),
            "min" => Math.Min(args[0], args[1]),
            "mod" => Math.Sqrt(args[0] * args[0] + args[1] * args[1] + args[2] * args[2]),
            "pin" => args[1] < args[0] ? args[0] : args[1] > args[2] ? args[2] : args[1],
            "sat2" => args[0] * Math.Sin(Math.Atan2(args[2], args[1])),
            "sin" => args[0] * Math.Sin(AngleRadians(args[1])),
            "sqrt" => args[0] < 0 ? throw new FormulaException($"{label} takes the square root of a negative value") : Math.Sqrt(args[0]),
            "tan" => args[0] * Math.Tan(AngleRadians(args[1])),
            "val" => args[0],
            _ => throw new FormulaException($"{label} uses an unsupported operator"),
        };
        if (!double.IsFinite(result) || Math.Abs(result) > MaxCoordinate)
            throw new FormulaException($"{label} evaluates outside the DrawingML signed 32-bit range");
        return result == -0d ? 0d : result;
    }

    private static double Operand(string token, IReadOnlyDictionary<string, double> values) =>
        IntegerRegex().IsMatch(token)
            ? long.Parse(token, NumberStyles.Integer, CultureInfo.InvariantCulture)
            : values[token];

    private static double AngleRadians(double value) => value / AngleUnitsPerDegree * Math.PI / 180d;

    private static Dictionary<string, double> Builtins(long widthEmu, long heightEmu)
    {
        if (widthEmu is <= 0 or > MaxCoordinate || heightEmu is <= 0 or > MaxCoordinate)
            throw new FormulaException("formula evaluation requires positive signed-32-bit shape extents");
        var w = (double)widthEmu;
        var h = (double)heightEmu;
        var ss = Math.Min(w, h);
        return new Dictionary<string, double>(StringComparer.Ordinal)
        {
            ["3cd4"] = 16_200_000, ["3cd8"] = 8_100_000, ["5cd8"] = 13_500_000, ["7cd8"] = 18_900_000,
            ["b"] = h, ["cd2"] = 10_800_000, ["cd4"] = 5_400_000, ["cd8"] = 2_700_000, ["h"] = h,
            ["hc"] = w / 2, ["hd2"] = h / 2, ["hd3"] = h / 3, ["hd4"] = h / 4, ["hd5"] = h / 5,
            ["hd6"] = h / 6, ["hd8"] = h / 8, ["l"] = 0, ["ls"] = Math.Max(w, h), ["r"] = w, ["ss"] = ss,
            ["ssd2"] = ss / 2, ["ssd4"] = ss / 4, ["ssd6"] = ss / 6, ["ssd8"] = ss / 8,
            ["ssd16"] = ss / 16, ["ssd32"] = ss / 32, ["t"] = 0, ["vc"] = h / 2, ["w"] = w,
            ["wd2"] = w / 2, ["wd3"] = w / 3, ["wd4"] = w / 4, ["wd5"] = w / 5, ["wd6"] = w / 6,
            ["wd8"] = w / 8, ["wd10"] = w / 10,
        };
    }

    private static bool HasUnsupportedShapeGuide(A.ShapeGuide guide) =>
        !string.IsNullOrEmpty(guide.InnerXml) ||
        guide.GetAttributes().Any(attribute => !string.IsNullOrEmpty(attribute.NamespaceUri) || attribute.LocalName is not ("name" or "fmla"));

    [GeneratedRegex("^[A-Za-z_][A-Za-z0-9_.-]{0,127}$", RegexOptions.CultureInvariant)]
    private static partial Regex GuideNameRegex();

    [GeneratedRegex("^-?\\d+$", RegexOptions.CultureInvariant)]
    private static partial Regex IntegerRegex();
}
