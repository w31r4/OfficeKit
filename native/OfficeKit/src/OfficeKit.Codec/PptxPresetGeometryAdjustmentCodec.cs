using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical ordered PPJ projection for DrawingML preset-shape adjustments.
// The public language owns values only. Native guide names are fixed by the
// ECMA-376 preset definitions and stay inside this codec boundary.
internal static class PptxPresetGeometryAdjustmentCodec
{
    internal const int MinimumValue = -21_600_000;
    internal const int MaximumValue = 21_600_000;

    private static readonly IReadOnlyDictionary<string, string[]> Profiles =
        new Dictionary<string, string[]>(StringComparer.Ordinal)
        {
            ["rect"] = [],
            ["textbox"] = [],
            ["line"] = [],
            ["roundRect"] = ["adj"],
            ["ellipse"] = [],
            ["triangle"] = ["adj"],
            ["rightTriangle"] = [],
            ["diamond"] = [],
            ["parallelogram"] = ["adj"],
            ["trapezoid"] = ["adj"],
            ["pentagon"] = ["hf", "vf"],
            ["hexagon"] = ["adj", "vf"],
            ["heptagon"] = ["hf", "vf"],
            ["octagon"] = ["adj"],
            ["chevron"] = ["adj"],
            ["homePlate"] = ["adj"],
            ["pie"] = ["adj1", "adj2"],
            ["arc"] = ["adj1", "adj2"],
            ["donut"] = ["adj"],
            ["blockArc"] = ["adj1", "adj2", "adj3"],
            ["heart"] = [],
            ["lightningBolt"] = [],
            ["sun"] = ["adj"],
            ["moon"] = ["adj"],
            ["cloud"] = [],
            ["star4"] = ["adj"],
            ["star5"] = ["adj", "hf", "vf"],
            ["star6"] = ["adj", "hf"],
            ["star8"] = ["adj"],
            ["star10"] = ["adj", "hf"],
            ["star12"] = ["adj"],
            ["leftArrow"] = ["adj1", "adj2"],
            ["rightArrow"] = ["adj1", "adj2"],
            ["upArrow"] = ["adj1", "adj2"],
            ["downArrow"] = ["adj1", "adj2"],
            ["leftRightArrow"] = ["adj1", "adj2"],
            ["upDownArrow"] = ["adj1", "adj2"],
            ["quadArrow"] = ["adj1", "adj2", "adj3"],
            ["bentArrow"] = ["adj1", "adj2", "adj3", "adj4"],
            ["uturnArrow"] = ["adj1", "adj2", "adj3", "adj4", "adj5"],
            ["circularArrow"] = ["adj1", "adj2", "adj3", "adj4", "adj5"],
            ["wedgeRoundRectCallout"] = ["adj1", "adj2", "adj3"],
            ["wedgeEllipseCallout"] = ["adj1", "adj2"],
            ["bracePair"] = ["adj"],
            ["bracketPair"] = ["adj"],
            ["flowChartProcess"] = [],
            ["flowChartDecision"] = [],
            ["flowChartData"] = [],
            ["flowChartTerminator"] = [],
            ["flowChartDocument"] = [],
            ["flowChartPreparation"] = [],
        };

    internal static bool HasProfile(string geometry) => Profiles.ContainsKey(geometry);

    internal static bool TryRead(A.PresetGeometry? native, string geometry, out int[] values)
    {
        values = [];
        if (native is null || !Profiles.TryGetValue(geometry, out var names) ||
            !PptxCustomGeometryCodec.TryPreset(geometry, out var expectedPreset) ||
            native.Preset?.Value is not { } actualPreset || !actualPreset.Equals(expectedPreset) ||
            !HasOnlyAttributes(native, "prst") || native.ChildElements.Count != 1 ||
            native.FirstChild is not A.AdjustValueList list || list.HasAttributes)
            return false;

        var guides = list.Elements<A.ShapeGuide>().ToArray();
        if (list.ChildElements.Count != guides.Length) return false;
        if (guides.Length == 0) return true;
        if (guides.Length != names.Length) return false;

        values = new int[guides.Length];
        for (var index = 0; index < guides.Length; index++)
        {
            var guide = guides[index];
            var formula = guide.Formula?.Value;
            if (!HasOnlyAttributes(guide, "name", "fmla") || guide.ChildElements.Count != 0 ||
                !string.Equals(guide.Name?.Value, names[index], StringComparison.Ordinal) ||
                formula is null || !formula.StartsWith("val ", StringComparison.Ordinal) ||
                !int.TryParse(formula.AsSpan(4), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var value) ||
                value is < MinimumValue or > MaximumValue)
                return false;
            values[index] = value;
        }
        return true;
    }

    internal static void Read(A.PresetGeometry? native, string geometry, PresentationShape target)
    {
        if (TryRead(native, geometry, out var values)) target.PresetAdjustments.Add(values);
    }

    internal static void Validate(string geometry, IEnumerable<int> values, string shapeId)
    {
        if (!Profiles.TryGetValue(geometry, out var names))
            throw new CodecException("unsupported_presentation_geometry", $"Presentation shape {shapeId} uses unsupported preset geometry {geometry}.");
        var materialized = values as IReadOnlyCollection<int> ?? values.ToArray();
        if (materialized.Count != 0 && materialized.Count != names.Length)
            throw new CodecException(
                "invalid_presentation_geometry",
                $"Presentation shape {shapeId} preset geometry {geometry} requires either no explicit adjustments or exactly {names.Length} ordered values.");
        if (materialized.Any(value => value is < MinimumValue or > MaximumValue))
            throw new CodecException(
                "invalid_presentation_geometry",
                $"Presentation shape {shapeId} preset adjustments must be between {MinimumValue} and {MaximumValue}.");
    }

    internal static void Apply(A.PresetGeometry native, string geometry, IEnumerable<int> values, string shapeId)
    {
        var materialized = values.ToArray();
        Validate(geometry, materialized, shapeId);
        var names = Profiles[geometry];
        native.RemoveAllChildren();
        native.Append(new A.AdjustValueList(materialized.Select((value, index) => new A.ShapeGuide
        {
            Name = names[index],
            Formula = $"val {value.ToString(CultureInfo.InvariantCulture)}",
        })));
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var attributes = element.GetAttributes();
        return attributes.Count == names.Length && attributes.All(attribute =>
            attribute.NamespaceUri.Length == 0 && names.Contains(attribute.LocalName, StringComparer.Ordinal));
    }
}
