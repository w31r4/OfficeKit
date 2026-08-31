using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PpjSmartArtDefinition(string LayoutProfile, JsonElement Root);

internal static partial class PpjSmartArtDefinitionCodec
{
    private const int MaxDefinitionBytes = 1024 * 1024;
    private const string Schema = "office-kit/smartart-definition/v1";
    private static readonly Lazy<IReadOnlySet<string>> Profiles = new(LoadProfiles, LazyThreadSafetyMode.ExecutionAndPublication);

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
            if (!Profiles.Value.Contains(profile))
                throw Invalid(path + ".layout.profile", $"unsupported operator profile {profile}");
            if (layout.TryGetProperty("operators", out var operators))
            {
                if (operators.ValueKind != JsonValueKind.Array)
                    throw Invalid(path + ".layout.operators", "operators must be an array");
                if (operators.GetArrayLength() > 0)
                    throw new CodecException(
                        "unsupported_ppj_compile_feature",
                        $"PPJ {path}.layout.operators: custom SmartArt operators are preserved by the asset contract but are not executable by this engine slice.",
                        path + ".layout.operators");
            }
            ValidateSection(root.GetProperty("style"), path + ".style");
            ValidateSection(root.GetProperty("colors"), path + ".colors");
            return new(profile, root.Clone());
        }
    }

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

    private static IReadOnlySet<string> LoadProfiles()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OfficeKit.Ppj.SmartArt.OperatorProfiles.json")
            ?? throw new InvalidOperationException("Embedded SmartArt operator profiles are missing.");
        using var document = JsonDocument.Parse(stream);
        if (document.RootElement.GetProperty("schema").GetString() != "office-kit/smartart-operator-profiles/v1")
            throw new InvalidOperationException("Embedded SmartArt operator profile manifest has the wrong schema.");
        return document.RootElement.GetProperty("profiles").EnumerateArray()
            .Select(profile => profile.GetProperty("id").GetString()!)
            .ToHashSet(StringComparer.Ordinal);
    }

    private static CodecException Invalid(string path, string detail) =>
        new("ppj.smartArt.definitionInvalid", $"SmartArt {detail}.", path);

    [GeneratedRegex("^[A-Za-z][A-Za-z0-9._-]{0,127}$", RegexOptions.CultureInvariant)]
    private static partial Regex IdPattern();
}
