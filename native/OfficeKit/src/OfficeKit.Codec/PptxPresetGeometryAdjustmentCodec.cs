using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical ordered PPJ projection for DrawingML preset-shape adjustments.
// The public language owns values only. Native guide names are fixed by the
// ECMA-376 preset definitions and stay inside this codec boundary.
internal static class PptxPresetGeometryAdjustmentCodec
{
    private const string ResourceName = "OfficeKit.Ppj.PresetGeometryProfiles.json";
    private sealed record Profile(string NativeToken, string[] Guides, bool Alias);

    internal static readonly int MinimumValue;
    internal static readonly int MaximumValue;
    private static readonly IReadOnlyDictionary<string, Profile> Profiles;
    private static readonly IReadOnlyDictionary<string, string> PreferredNames;

    static PptxPresetGeometryAdjustmentCodec()
    {
        using var stream = typeof(PptxPresetGeometryAdjustmentCodec).Assembly.GetManifestResourceStream(ResourceName) ??
            throw new InvalidOperationException($"Embedded preset geometry profile resource {ResourceName} is missing.");
        using var document = JsonDocument.Parse(stream);
        var root = document.RootElement;
        if (root.GetProperty("schema").GetString() != "office-kit/ppj-preset-geometry-profiles/v1")
            throw new InvalidOperationException("Embedded preset geometry profile schema is invalid.");
        MinimumValue = root.GetProperty("minimumValue").GetInt32();
        MaximumValue = root.GetProperty("maximumValue").GetInt32();
        Profiles = root.GetProperty("profiles").EnumerateObject().ToDictionary(
            property => property.Name,
            property => new Profile(
                property.Value.GetProperty("nativeToken").GetString()!,
                property.Value.GetProperty("guides").EnumerateArray().Select(item => item.GetString()!).ToArray(),
                property.Value.TryGetProperty("alias", out var alias) && alias.GetBoolean()),
            StringComparer.Ordinal);
        PreferredNames = Profiles
            .Where(entry => !entry.Value.Alias)
            .ToDictionary(entry => entry.Value.NativeToken, entry => entry.Key, StringComparer.Ordinal);
    }

    private static bool TryProfile(string geometry, out Profile profile)
    {
        if (geometry is "textbox" or "line")
        {
            profile = new Profile(geometry == "textbox" ? "rect" : "line", [], false);
            return true;
        }
        return Profiles.TryGetValue(geometry, out profile!);
    }

    internal static bool HasProfile(string geometry) => TryProfile(geometry, out _);

    internal static bool TryPreset(string geometry, out A.ShapeTypeValues preset)
    {
        if (TryProfile(geometry, out var profile))
        {
            preset = new A.ShapeTypeValues(profile.NativeToken);
            return true;
        }
        preset = default;
        return false;
    }

    internal static bool TryPresetName(A.ShapeTypeValues preset, out string name) =>
        PreferredNames.TryGetValue(preset.ToString(), out name!);

    internal static bool TryExpectedCount(string geometry, out int count)
    {
        if (TryProfile(geometry, out var profile))
        {
            count = profile.Guides.Length;
            return true;
        }
        count = 0;
        return false;
    }

    internal static bool TryRead(A.PresetGeometry? native, string geometry, out int[] values)
    {
        values = [];
        if (native is null || !TryProfile(geometry, out var profile) ||
            !TryPreset(geometry, out var expectedPreset) ||
            native.Preset?.Value is not { } actualPreset || !actualPreset.Equals(expectedPreset) ||
            !HasOnlyAttributes(native, "prst") || native.ChildElements.Count != 1 ||
            native.FirstChild is not A.AdjustValueList list || list.HasAttributes)
            return false;

        var guides = list.Elements<A.ShapeGuide>().ToArray();
        if (list.ChildElements.Count != guides.Length) return false;
        if (guides.Length == 0) return true;
        if (guides.Length != profile.Guides.Length) return false;

        values = new int[guides.Length];
        for (var index = 0; index < guides.Length; index++)
        {
            var guide = guides[index];
            var formula = guide.Formula?.Value;
            if (!HasOnlyAttributes(guide, "name", "fmla") || guide.ChildElements.Count != 0 ||
                !string.Equals(guide.Name?.Value, profile.Guides[index], StringComparison.Ordinal) ||
                formula is null || !formula.StartsWith("val ", StringComparison.Ordinal) ||
                !int.TryParse(formula.AsSpan(4), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var value) ||
                value < MinimumValue || value > MaximumValue)
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
        if (!TryProfile(geometry, out var profile))
            throw new CodecException("unsupported_presentation_geometry", $"Presentation shape {shapeId} uses unsupported preset geometry {geometry}.");
        var materialized = values as IReadOnlyCollection<int> ?? values.ToArray();
        if (materialized.Count != 0 && materialized.Count != profile.Guides.Length)
            throw new CodecException(
                "invalid_presentation_geometry",
                $"Presentation shape {shapeId} preset geometry {geometry} requires either no explicit adjustments or exactly {profile.Guides.Length} ordered values.");
        if (materialized.Any(value => value < MinimumValue || value > MaximumValue))
            throw new CodecException(
                "invalid_presentation_geometry",
                $"Presentation shape {shapeId} preset adjustments must be between {MinimumValue} and {MaximumValue}.");
    }

    internal static void Apply(A.PresetGeometry native, string geometry, IEnumerable<int> values, string shapeId)
    {
        var materialized = values.ToArray();
        Validate(geometry, materialized, shapeId);
        _ = TryProfile(geometry, out var profile);
        native.RemoveAllChildren();
        native.Append(new A.AdjustValueList(materialized.Select((value, index) => new A.ShapeGuide
        {
            Name = profile.Guides[index],
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
