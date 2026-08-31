using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PpjSmartArtExecutionPlan(
    string Placement,
    int? Columns = null,
    double? GapPoints = null,
    bool Reverse = false);

internal sealed record PpjSmartArtDefinition(
    string LayoutProfile,
    PpjSmartArtExecutionPlan Execution,
    JsonElement Root);

internal static partial class PpjSmartArtDefinitionCodec
{
    private const int MaxDefinitionBytes = 1024 * 1024;
    private const string Schema = "office-kit/smartart-definition/v1";
    private static readonly Lazy<IReadOnlyDictionary<string, PpjSmartArtExecutionPlan>> Profiles =
        new(LoadProfiles, LazyThreadSafetyMode.ExecutionAndPublication);

    internal static PpjSmartArtDefinition BuiltIn(string profile)
    {
        if (!Profiles.Value.TryGetValue(profile, out var execution))
            throw new CodecException(
                "unsupported_ppj_compile_feature",
                $"PPJ authored SmartArt layout {profile} is not compiler-owned.");
        return new(profile, execution, default);
    }

    internal static PpjSmartArtDefinition Parse(Asset asset, string path)
    {
        if (asset.Data.IsEmpty || asset.Data.Length > MaxDefinitionBytes)
            throw Invalid(path, $"definition bytes must contain 1 through {MaxDefinitionBytes} bytes");
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(asset.Data.Memory, new JsonDocumentOptions
            {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = 64,
            });
        }
        catch (JsonException exception)
        {
            throw Invalid(path, $"definition is not strict JSON: {exception.Message}");
        }
        using (document)
        {
            var root = document.RootElement;
            RequireObject(root, path, ["schema", "layout", "style", "colors"], ["schema", "layout", "style", "colors"]);
            if (root.GetProperty("schema").ValueKind != JsonValueKind.String || root.GetProperty("schema").GetString() != Schema)
                throw Invalid(path + ".schema", $"value must equal {Schema}");
            var layout = root.GetProperty("layout");
            RequireObject(layout, path + ".layout", ["id", "profile", "operators"], ["id", "profile"]);
            ValidateId(layout.GetProperty("id"), path + ".layout.id");
            var profileValue = layout.GetProperty("profile");
            var profile = profileValue.ValueKind == JsonValueKind.String ? profileValue.GetString()! : string.Empty;
            if (!Profiles.Value.TryGetValue(profile, out var execution))
                throw Invalid(path + ".layout.profile", $"unsupported operator profile {profile}");
            if (layout.TryGetProperty("operators", out var operators))
            {
                if (operators.ValueKind != JsonValueKind.Array)
                    throw Invalid(path + ".layout.operators", "operators must be an array");
                if (operators.GetArrayLength() > 256)
                    throw Invalid(path + ".layout.operators", "operators cannot contain more than 256 entries");
                execution = ParseOperators(operators, execution, path + ".layout.operators");
            }
            ValidateSection(root.GetProperty("style"), path + ".style");
            ValidateSection(root.GetProperty("colors"), path + ".colors");
            return new(profile, execution, root.Clone());
        }
    }

    private static PpjSmartArtExecutionPlan ParseOperators(
        JsonElement operators,
        PpjSmartArtExecutionPlan inherited,
        string path)
    {
        var placement = inherited.Placement;
        var columns = inherited.Columns;
        var gapPoints = inherited.GapPoints;
        var reverse = inherited.Reverse;
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var index = 0;
        foreach (var value in operators.EnumerateArray())
        {
            var operatorPath = $"{path}[{index}]";
            RequireObject(value, operatorPath, ["id", "kind", "input", "arguments"], ["id", "kind"]);
            ValidateId(value.GetProperty("id"), operatorPath + ".id");
            if (!ids.Add(value.GetProperty("id").GetString()!))
                throw Invalid(operatorPath + ".id", "operator identifiers must be unique");
            if (value.TryGetProperty("input", out var input) &&
                (input.ValueKind != JsonValueKind.String || input.GetString() != "nodes"))
                throw Unsupported(operatorPath + ".input", "only the nodes input is executable");
            var kindValue = value.GetProperty("kind");
            if (kindValue.ValueKind != JsonValueKind.String)
                throw Invalid(operatorPath + ".kind", "operator kind must be a string");
            var kind = kindValue.GetString()!;
            if (kind is not ("algorithm" or "constraint" or "rule"))
                throw Unsupported(operatorPath + ".kind", $"operator kind {kind} is not executable");
            if (!value.TryGetProperty("arguments", out var arguments))
                throw Invalid(operatorPath + ".arguments", $"{kind} operators require arguments");
            if (arguments.ValueKind != JsonValueKind.Object)
                throw Invalid(operatorPath + ".arguments", "operator arguments must be an object");
            if (arguments.EnumerateObject().Count() > 32)
                throw Invalid(operatorPath + ".arguments", "operator arguments cannot contain more than 32 properties");

            switch (kind)
            {
                case "algorithm":
                    RequireArguments(arguments, operatorPath, ["placement"], ["placement"]);
                    var placementValue = arguments.GetProperty("placement");
                    if (placementValue.ValueKind != JsonValueKind.String ||
                        !SupportedPlacements.Contains(placementValue.GetString()!))
                        throw Unsupported(operatorPath + ".arguments.placement", "placement is not in the bounded SmartArt placement catalog");
                    placement = placementValue.GetString()!;
                    break;
                case "constraint":
                    RequireArguments(arguments, operatorPath, ["gapPoints"], ["gapPoints"]);
                    var gapValue = arguments.GetProperty("gapPoints");
                    if (gapValue.ValueKind != JsonValueKind.Number || !gapValue.TryGetDouble(out var gap) ||
                        !double.IsFinite(gap) || gap is < 0 or > 72)
                        throw Invalid(operatorPath + ".arguments.gapPoints", "gapPoints must be a finite number from 0 through 72");
                    gapPoints = gap;
                    break;
                case "rule":
                    RequireArguments(arguments, operatorPath, ["columns", "reverse"], []);
                    if (!arguments.EnumerateObject().Any())
                        throw Invalid(operatorPath + ".arguments", "rule operators require columns or reverse");
                    if (arguments.TryGetProperty("columns", out var columnValue))
                    {
                        if (columnValue.ValueKind != JsonValueKind.Number || !columnValue.TryGetInt32(out var count) || count is < 1 or > 64)
                            throw Invalid(operatorPath + ".arguments.columns", "columns must be an integer from 1 through 64");
                        columns = count;
                    }
                    if (arguments.TryGetProperty("reverse", out var reverseValue))
                    {
                        if (reverseValue.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                            throw Invalid(operatorPath + ".arguments.reverse", "reverse must be a boolean");
                        reverse = reverseValue.GetBoolean();
                    }
                    break;
            }
            index++;
        }
        if (columns is not null && placement is not ("grid" or "horizontal-grid" or "square-grid" or "square-grid-picture"))
            throw Unsupported(path, $"columns cannot be applied to the {placement} placement");
        return new(placement, columns, gapPoints, reverse);
    }

    private static void RequireArguments(JsonElement value, string operatorPath, string[] allowed, string[] required) =>
        RequireObject(value, operatorPath + ".arguments", allowed, required);

    private static void ValidateSection(JsonElement section, string path)
    {
        RequireObject(section, path, ["id", "labels"], ["id"]);
        ValidateId(section.GetProperty("id"), path + ".id");
        if (!section.TryGetProperty("labels", out var labels)) return;
        if (labels.ValueKind != JsonValueKind.Array || labels.GetArrayLength() > 64 ||
            labels.EnumerateArray().Any(label => label.ValueKind != JsonValueKind.String || label.GetString()!.Length > 128))
            throw Invalid(path + ".labels", "labels must be an array of at most 64 strings with at most 128 characters each");
    }

    private static void RequireObject(JsonElement value, string path, string[] allowed, string[] required)
    {
        if (value.ValueKind != JsonValueKind.Object)
            throw Invalid(path, "value must be an object");
        var allowedNames = allowed.ToHashSet(StringComparer.Ordinal);
        foreach (var property in value.EnumerateObject())
            if (!allowedNames.Contains(property.Name))
                throw Invalid(path + "." + property.Name, "property is not part of the SmartArt definition v1 schema");
        foreach (var name in required)
            if (!value.TryGetProperty(name, out _)) throw Invalid(path + "." + name, "required property is missing");
    }

    private static void ValidateId(JsonElement value, string path)
    {
        if (value.ValueKind != JsonValueKind.String || !IdPattern().IsMatch(value.GetString()!))
            throw Invalid(path, "value must be a SmartArt definition identifier");
    }

    private static IReadOnlyDictionary<string, PpjSmartArtExecutionPlan> LoadProfiles()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OfficeKit.Ppj.SmartArt.OperatorProfiles.json")
            ?? throw new InvalidOperationException("Embedded SmartArt operator profiles are missing.");
        using var document = JsonDocument.Parse(stream);
        if (document.RootElement.GetProperty("schema").GetString() != "office-kit/smartart-operator-profiles/v1")
            throw new InvalidOperationException("Embedded SmartArt operator profile manifest has the wrong schema.");
        return document.RootElement.GetProperty("profiles").EnumerateArray()
            .ToDictionary(
                profile => profile.GetProperty("id").GetString()!,
                profile => new PpjSmartArtExecutionPlan(profile.GetProperty("placement").GetString()!),
                StringComparer.Ordinal);
    }

    private static CodecException Unsupported(string path, string detail) =>
        new("unsupported_ppj_compile_feature", $"PPJ {path}: SmartArt {detail}.", path);

    private static readonly IReadOnlySet<string> SupportedPlacements = new HashSet<string>(StringComparer.Ordinal)
    {
        "grid",
        "horizontal-grid",
        "radial",
        "depth-levels",
        "center-radial",
        "square-grid",
        "stacked-width",
        "square-grid-picture",
    };

    private static CodecException Invalid(string path, string detail) =>
        new("ppj.smartArt.definitionInvalid", $"SmartArt {detail}.", path);

    [GeneratedRegex("^[A-Za-z][A-Za-z0-9._-]{0,127}$", RegexOptions.CultureInvariant)]
    private static partial Regex IdPattern();
}
