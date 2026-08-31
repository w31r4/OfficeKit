using System.Buffers;
using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeKit.Codec;

internal sealed record PpjDiagnostic(string Code, string Message, string Path);

internal sealed record PpjValidationResult(
    PpjProgramModel? Program,
    IReadOnlyList<PpjDiagnostic> Diagnostics,
    byte[] CanonicalJson,
    string ProgramSha256,
    PpjExpansionResult? Expansion) : IDisposable
{
    internal JsonDocument? Document { get; init; }
    internal bool IsValid => Program is not null && Diagnostics.Count == 0;
    public void Dispose() => Document?.Dispose();
}

internal static class PpjProgramValidator
{
    internal const int MaxSourceBytes = 16 * 1024 * 1024;
    internal const int MaxPages = 512;
    internal const int MaxExpandedElements = 100_000;
    internal const int MaxRepeatItems = 1_024;
    internal const int MaxComponentDepth = 16;
    private const int MaxJsonDepth = 96;

    internal static PpjValidationResult Validate(ReadOnlyMemory<byte> json)
    {
        var diagnostics = new List<PpjDiagnostic>();
        if (json.Length == 0)
        {
            diagnostics.Add(new("ppj.empty", "PPJ input must not be empty.", "$"));
            return Invalid(diagnostics);
        }
        if (json.Length > MaxSourceBytes)
        {
            diagnostics.Add(new(
                "ppj.sourceBudget",
                $"PPJ input has {json.Length} bytes and exceeds the {MaxSourceBytes}-byte source budget.",
                "$"));
            return Invalid(diagnostics);
        }
        var jsonSpan = json.Span;
        if (json.Length >= 3 && jsonSpan[0] == 0xef && jsonSpan[1] == 0xbb && jsonSpan[2] == 0xbf)
        {
            diagnostics.Add(new("ppj.utf8Bom", "PPJ must be UTF-8 JSON without a byte-order mark.", "$"));
            return Invalid(diagnostics);
        }

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json, new JsonDocumentOptions
            {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = MaxJsonDepth,
            });
        }
        catch (JsonException exception)
        {
            diagnostics.Add(new(
                "ppj.invalidJson",
                $"PPJ is not strict UTF-8 JSON: {exception.Message}",
                exception.Path ?? "$"));
            return Invalid(diagnostics);
        }
        var root = document.RootElement;

        FindDuplicateProperties(root, new StringBuilder("$"), [], 0, diagnostics);
        if (diagnostics.Count == 0)
            PpjJsonSchemaValidator.Validate(root, diagnostics);
        if (diagnostics.Count != 0)
            return Invalid(diagnostics, document);

        PpjProgramModel program;
        try
        {
            program = PpjProgramParser.Parse(root);
        }
        catch (Exception exception) when (exception is KeyNotFoundException or InvalidOperationException or FormatException)
        {
            diagnostics.Add(new("ppj.modelProjection", "Validated PPJ could not be projected into the native typed model.", "$"));
            return Invalid(diagnostics, document);
        }

        PpjSemanticValidator.Validate(program, diagnostics);
        if (diagnostics.Count != 0)
            return Invalid(diagnostics, document);

        var canonical = PpjCanonicalJson.Write(root);
        var hash = Convert.ToHexString(SHA256.HashData(canonical)).ToLowerInvariant();
        var expansion = PpjComponentExpander.Expand(program, hash, diagnostics);
        if (diagnostics.Count != 0 || expansion is null)
            return Invalid(diagnostics, document);
        return new PpjValidationResult(program, diagnostics, canonical, hash, expansion) { Document = document };
    }

    private static PpjValidationResult Invalid(IReadOnlyList<PpjDiagnostic> diagnostics, JsonDocument? document = null)
    {
        document?.Dispose();
        return new(null, diagnostics, [], string.Empty, null);
    }

    private static void FindDuplicateProperties(
        JsonElement value,
        StringBuilder path,
        List<HashSet<string>> propertySets,
        int depth,
        List<PpjDiagnostic> diagnostics)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            while (propertySets.Count <= depth)
                propertySets.Add(new HashSet<string>(StringComparer.Ordinal));
            var names = propertySets[depth];
            names.Clear();
            foreach (var property in value.EnumerateObject())
            {
                var propertyName = property.Name;
                var length = AppendJsonProperty(path, propertyName);
                if (!names.Add(propertyName))
                    diagnostics.Add(new("ppj.duplicateProperty", $"Property {propertyName} appears more than once.", path.ToString()));
                FindDuplicateProperties(property.Value, path, propertySets, depth + 1, diagnostics);
                path.Length = length;
            }
            names.Clear();
        }
        else if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                var length = path.Length;
                path.Append('[').Append(index++).Append(']');
                FindDuplicateProperties(item, path, propertySets, depth + 1, diagnostics);
                path.Length = length;
            }
        }
    }

    private static int AppendJsonProperty(StringBuilder path, string property)
    {
        var length = path.Length;
        if (property.All(character => char.IsAsciiLetterOrDigit(character) || character is '_' or '-'))
            path.Append('.').Append(property);
        else
            path.Append("['").Append(property.Replace("'", "\\'", StringComparison.Ordinal)).Append("']");
        return length;
    }

}

internal static class PpjJsonSchemaValidator
{
    private static readonly Lazy<JsonElement> Schema = new(LoadSchema, LazyThreadSafetyMode.ExecutionAndPublication);
    private static readonly Lazy<IReadOnlyDictionary<string, JsonElement>> Definitions =
        new(LoadDefinitions, LazyThreadSafetyMode.ExecutionAndPublication);
    private static readonly ConcurrentDictionary<JsonElement, ObjectSchemaMetadata> ObjectMetadata = new();
    private static readonly ConcurrentDictionary<JsonElement, ChoiceMetadata> ChoiceMetadataCache = new();
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(100);
    private static readonly string[] ChoiceDiscriminators = ["type", "kind", "mode"];

    private sealed record ObjectSchemaMetadata(
        IReadOnlyList<string> Required,
        IReadOnlyDictionary<string, JsonElement> Declared,
        IReadOnlySet<string>? Evaluated);

    private sealed record ChoiceMetadata(
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, JsonElement>> Discriminated,
        IReadOnlyList<(JsonElement Choice, JsonElement Type)>? Typed);

    internal static void Validate(JsonElement instance, List<PpjDiagnostic> diagnostics)
    {
        var path = new StringBuilder("$");
        ValidateNode(instance, Schema.Value, path, diagnostics);
    }

    private static JsonElement LoadSchema()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OfficeKit.Ppj.V1.Schema.json")
            ?? throw new InvalidOperationException("Embedded PPJ v1 schema is missing.");
        using var document = JsonDocument.Parse(stream);
        return document.RootElement.Clone();
    }

    private static IReadOnlyDictionary<string, JsonElement> LoadDefinitions() =>
        Schema.Value.GetProperty("$defs").EnumerateObject()
            .ToDictionary(property => property.Name, property => property.Value, StringComparer.Ordinal);

    private static void ValidateNode(JsonElement instance, JsonElement schema, StringBuilder path, List<PpjDiagnostic> diagnostics)
    {
        if (schema.ValueKind == JsonValueKind.True) return;
        if (schema.ValueKind == JsonValueKind.False)
        {
            diagnostics.Add(new("ppj.schema.false", "Value is prohibited by the PPJ schema.", path.ToString()));
            return;
        }
        if (schema.ValueKind != JsonValueKind.Object)
        {
            diagnostics.Add(new("ppj.schema.internal", "PPJ schema contains an invalid rule.", path.ToString()));
            return;
        }

        if (schema.TryGetProperty("$ref", out var reference))
        {
            ValidateNode(instance, ResolveReference(reference.GetString()!), path, diagnostics);
            return;
        }

        if (schema.TryGetProperty("allOf", out var allOf))
        {
            foreach (var child in allOf.EnumerateArray())
                ValidateNode(instance, child, path, diagnostics);
        }

        if (schema.TryGetProperty("oneOf", out var oneOf))
            ValidateChoice(instance, oneOf, path, diagnostics, requireExactlyOne: true);
        if (schema.TryGetProperty("anyOf", out var anyOf))
            ValidateChoice(instance, anyOf, path, diagnostics, requireExactlyOne: false);

        if (schema.TryGetProperty("type", out var type) && !MatchesType(instance, type))
        {
            diagnostics.Add(new(
                "ppj.schema.type",
                $"Expected {DescribeType(type)}, received {DescribeInstance(instance)}.",
                path.ToString()));
            return;
        }

        if (schema.TryGetProperty("const", out var constant) && !JsonEqual(instance, constant))
            diagnostics.Add(new("ppj.schema.const", $"Value must equal {constant.GetRawText()}.", path.ToString()));

        if (schema.TryGetProperty("enum", out var enumeration) && !enumeration.EnumerateArray().Any(item => JsonEqual(instance, item)))
            diagnostics.Add(new("ppj.schema.enum", "Value is not one of the allowed PPJ values.", path.ToString()));

        switch (instance.ValueKind)
        {
            case JsonValueKind.Object:
                ValidateObject(instance, schema, path, diagnostics);
                break;
            case JsonValueKind.Array:
                ValidateArray(instance, schema, path, diagnostics);
                break;
            case JsonValueKind.String:
                ValidateString(instance.GetString()!, schema, path, diagnostics);
                break;
            case JsonValueKind.Number:
                ValidateNumber(instance, schema, path, diagnostics);
                break;
        }
    }

    private static void ValidateChoice(
        JsonElement instance,
        JsonElement choices,
        StringBuilder path,
        List<PpjDiagnostic> diagnostics,
        bool requireExactlyOne)
    {
        if (TrySelectExclusiveChoice(instance, choices, out var selected))
        {
            var diagnosticCount = diagnostics.Count;
            ValidateNode(instance, selected, path, diagnostics);
            if (diagnostics.Count == diagnosticCount) return;
            diagnostics.RemoveRange(diagnosticCount, diagnostics.Count - diagnosticCount);
        }

        var branches = new List<List<PpjDiagnostic>>();
        foreach (var choice in choices.EnumerateArray())
        {
            var branch = new List<PpjDiagnostic>();
            ValidateNode(instance, choice, path, branch);
            branches.Add(branch);
        }
        var matches = branches.Count(branch => branch.Count == 0);
        if ((requireExactlyOne && matches == 1) || (!requireExactlyOne && matches > 0)) return;

        if (matches == 0 && branches.Count > 0)
            diagnostics.AddRange(branches.OrderBy(branch => branch.Count).ThenBy(branch => string.Join('\n', branch.Select(item => item.Code))).First());
        diagnostics.Add(new(
            requireExactlyOne ? "ppj.schema.oneOf" : "ppj.schema.anyOf",
            requireExactlyOne
                ? $"Value must match exactly one typed PPJ alternative; matched {matches}."
                : "Value must match at least one typed PPJ alternative.",
            path.ToString()));
    }

    private static bool TrySelectExclusiveChoice(JsonElement instance, JsonElement choices, out JsonElement selected)
    {
        selected = default;
        if (instance.ValueKind == JsonValueKind.Object)
        {
            foreach (var discriminator in ChoiceDiscriminators)
            {
                if (!instance.TryGetProperty(discriminator, out var value) || value.ValueKind != JsonValueKind.String)
                    continue;
                if (TrySelectDiscriminatedBranch(choices, discriminator, value.GetString()!, out selected))
                    return true;
            }
        }

        return TrySelectTypedBranch(instance, choices, out selected);
    }

    private static bool TrySelectDiscriminatedBranch(
        JsonElement choices,
        string discriminator,
        string value,
        out JsonElement selected)
    {
        selected = default;
        return ChoiceMetadataCache.GetOrAdd(choices, BuildChoiceMetadata)
            .Discriminated.TryGetValue(discriminator, out var branches) &&
            branches.TryGetValue(value, out selected);
    }

    private static bool TrySelectTypedBranch(JsonElement instance, JsonElement choices, out JsonElement selected)
    {
        selected = default;
        var matches = 0;
        var typed = ChoiceMetadataCache.GetOrAdd(choices, BuildChoiceMetadata).Typed;
        if (typed is null) return false;
        foreach (var (choice, type) in typed)
        {
            if (!MatchesType(instance, type)) continue;
            selected = choice;
            matches++;
        }
        return matches == 1;
    }

    private static ChoiceMetadata BuildChoiceMetadata(JsonElement choices)
    {
        var choiceList = choices.EnumerateArray().ToArray();
        var discriminated = new Dictionary<string, IReadOnlyDictionary<string, JsonElement>>(StringComparer.Ordinal);
        foreach (var discriminator in ChoiceDiscriminators)
        {
            var branches = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
            var complete = true;
            foreach (var choice in choiceList)
            {
                if (!RequiresProperty(choice, discriminator) ||
                    !TryFindStringConstProperty(choice, discriminator, out var expected) ||
                    !branches.TryAdd(expected, choice))
                {
                    complete = false;
                    break;
                }
            }
            if (complete) discriminated.Add(discriminator, branches);
        }

        List<(JsonElement Choice, JsonElement Type)>? typed = new(choiceList.Length);
        foreach (var choice in choiceList)
        {
            if (!TryFindType(choice, out var type))
            {
                typed = null;
                break;
            }
            typed.Add((choice, type));
        }
        return new(discriminated, typed);
    }

    private static bool RequiresProperty(JsonElement schema, string property)
    {
        if (schema.ValueKind != JsonValueKind.Object) return false;
        if (schema.TryGetProperty("$ref", out var reference))
            return RequiresProperty(ResolveReference(reference.GetString()!), property);
        if (schema.TryGetProperty("required", out var required) &&
            required.EnumerateArray().Any(item => property.Equals(item.GetString(), StringComparison.Ordinal)))
            return true;
        return schema.TryGetProperty("allOf", out var allOf) &&
               allOf.EnumerateArray().Any(child => RequiresProperty(child, property));
    }

    private static bool TryFindStringConstProperty(JsonElement schema, string property, out string value)
    {
        value = string.Empty;
        if (schema.ValueKind != JsonValueKind.Object) return false;
        if (schema.TryGetProperty("$ref", out var reference))
            return TryFindStringConstProperty(ResolveReference(reference.GetString()!), property, out value);
        if (schema.TryGetProperty("properties", out var properties) &&
            properties.TryGetProperty(property, out var propertySchema) &&
            propertySchema.TryGetProperty("const", out var constant) &&
            constant.ValueKind == JsonValueKind.String)
        {
            value = constant.GetString()!;
            return true;
        }
        if (!schema.TryGetProperty("allOf", out var allOf)) return false;
        foreach (var child in allOf.EnumerateArray())
        {
            if (TryFindStringConstProperty(child, property, out value)) return true;
        }
        return false;
    }

    private static bool TryFindType(JsonElement schema, out JsonElement type)
    {
        type = default;
        if (schema.ValueKind != JsonValueKind.Object) return false;
        if (schema.TryGetProperty("$ref", out var reference))
            return TryFindType(ResolveReference(reference.GetString()!), out type);
        if (schema.TryGetProperty("type", out type)) return true;
        if (!schema.TryGetProperty("allOf", out var allOf)) return false;
        foreach (var child in allOf.EnumerateArray())
        {
            if (TryFindType(child, out type)) return true;
        }
        return false;
    }

    private static void ValidateObject(JsonElement instance, JsonElement schema, StringBuilder path, List<PpjDiagnostic> diagnostics)
    {
        var metadata = ObjectMetadata.GetOrAdd(schema, BuildObjectMetadata);
        var propertyCount = instance.EnumerateObject().Count();
        if (schema.TryGetProperty("maxProperties", out var maximum) && propertyCount > maximum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"Object has {propertyCount} properties; maximum is {maximum.GetInt32()}.", path.ToString()));

        foreach (var name in metadata.Required)
        {
            if (!instance.TryGetProperty(name, out _))
            {
                var length = AppendProperty(path, name);
                diagnostics.Add(new("ppj.schema.required", $"Required property {name} is missing.", path.ToString()));
                path.Length = length;
            }
        }

        foreach (var property in instance.EnumerateObject())
        {
            var length = AppendProperty(path, property.Name);
            if (metadata.Declared.TryGetValue(property.Name, out var propertySchema))
                ValidateNode(property.Value, propertySchema, path, diagnostics);
            if (schema.TryGetProperty("propertyNames", out var propertyNameSchema))
                ValidateNode(StringElement(property.Name), propertyNameSchema, path, diagnostics);
            path.Length = length;
        }

        if (schema.TryGetProperty("additionalProperties", out var additional))
        {
            foreach (var property in instance.EnumerateObject().Where(item => !metadata.Declared.ContainsKey(item.Name)))
            {
                var length = AppendProperty(path, property.Name);
                if (additional.ValueKind == JsonValueKind.False)
                    diagnostics.Add(new("ppj.schema.unknownField", $"Unknown property {property.Name} is not allowed.", path.ToString()));
                else if (additional.ValueKind == JsonValueKind.Object || additional.ValueKind is JsonValueKind.True or JsonValueKind.False)
                    ValidateNode(property.Value, additional, path, diagnostics);
                path.Length = length;
            }
        }

        if (schema.TryGetProperty("unevaluatedProperties", out var unevaluated) && unevaluated.ValueKind == JsonValueKind.False)
        {
            var evaluated = metadata.Evaluated!;
            foreach (var property in instance.EnumerateObject().Where(item => !evaluated.Contains(item.Name)))
            {
                var length = AppendProperty(path, property.Name);
                diagnostics.Add(new("ppj.schema.unknownField", $"Unknown property {property.Name} is not allowed.", path.ToString()));
                path.Length = length;
            }
        }
    }

    private static ObjectSchemaMetadata BuildObjectMetadata(JsonElement schema)
    {
        var required = schema.TryGetProperty("required", out var requiredSchema)
            ? requiredSchema.EnumerateArray().Select(item => item.GetString()!).ToArray()
            : [];
        var declared = schema.TryGetProperty("properties", out var properties)
            ? properties.EnumerateObject().ToDictionary(item => item.Name, item => item.Value, StringComparer.Ordinal)
            : new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        var evaluated = schema.TryGetProperty("unevaluatedProperties", out var unevaluated) &&
                        unevaluated.ValueKind == JsonValueKind.False
            ? CollectDeclaredProperties(schema)
            : null;
        return new(required, declared, evaluated);
    }

    private static JsonElement StringElement(string value)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
            writer.WriteStringValue(value);
        using var document = JsonDocument.Parse(buffer.ToArray());
        return document.RootElement.Clone();
    }

    private static void ValidateArray(JsonElement instance, JsonElement schema, StringBuilder path, List<PpjDiagnostic> diagnostics)
    {
        var itemCount = instance.GetArrayLength();
        if (schema.TryGetProperty("minItems", out var minimum) && itemCount < minimum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"Array has {itemCount} items; minimum is {minimum.GetInt32()}.", path.ToString()));
        if (schema.TryGetProperty("maxItems", out var maximum) && itemCount > maximum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"Array has {itemCount} items; maximum is {maximum.GetInt32()}.", path.ToString()));
        if (schema.TryGetProperty("uniqueItems", out var unique) && unique.GetBoolean())
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var index = 0;
            foreach (var item in instance.EnumerateArray())
            {
                var key = Convert.ToBase64String(PpjCanonicalJson.Write(item));
                if (!seen.Add(key))
                {
                    var length = AppendIndex(path, index);
                    diagnostics.Add(new("ppj.schema.unique", "Array item duplicates an earlier value.", path.ToString()));
                    path.Length = length;
                }
                index++;
            }
        }
        if (schema.TryGetProperty("items", out var itemSchema))
        {
            var index = 0;
            foreach (var item in instance.EnumerateArray())
            {
                var length = AppendIndex(path, index++);
                ValidateNode(item, itemSchema, path, diagnostics);
                path.Length = length;
            }
        }
    }

    private static void ValidateString(string value, JsonElement schema, StringBuilder path, List<PpjDiagnostic> diagnostics)
    {
        var length = value.EnumerateRunes().Count();
        if (schema.TryGetProperty("minLength", out var minimum) && length < minimum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"String has {length} characters; minimum is {minimum.GetInt32()}.", path.ToString()));
        if (schema.TryGetProperty("maxLength", out var maximum) && length > maximum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"String has {length} characters; maximum is {maximum.GetInt32()}.", path.ToString()));
        if (schema.TryGetProperty("pattern", out var pattern))
        {
            try
            {
                if (!Regex.IsMatch(value, pattern.GetString()!, RegexOptions.CultureInvariant, RegexTimeout))
                    diagnostics.Add(new("ppj.schema.pattern", "String does not match the required PPJ pattern.", path.ToString()));
            }
            catch (RegexMatchTimeoutException)
            {
                diagnostics.Add(new("ppj.schema.pattern", "String pattern validation exceeded its bounded time budget.", path.ToString()));
            }
        }
        if (schema.TryGetProperty("format", out var format) && !MatchesFormat(value, format.GetString()!))
            diagnostics.Add(new("ppj.schema.format", $"String is not a valid {format.GetString()} value.", path.ToString()));
    }

    private static void ValidateNumber(JsonElement instance, JsonElement schema, StringBuilder path, List<PpjDiagnostic> diagnostics)
    {
        if (!instance.TryGetDecimal(out var value))
        {
            var floating = instance.GetDouble();
            if (!double.IsFinite(floating))
                diagnostics.Add(new("ppj.schema.number", "Number must be finite.", path.ToString()));
            return;
        }
        if (schema.TryGetProperty("minimum", out var minimum) && value < minimum.GetDecimal())
            diagnostics.Add(new("ppj.schema.minimum", $"Number must be at least {minimum.GetRawText()}.", path.ToString()));
        if (schema.TryGetProperty("maximum", out var maximum) && value > maximum.GetDecimal())
            diagnostics.Add(new("ppj.schema.maximum", $"Number must be at most {maximum.GetRawText()}.", path.ToString()));
        if (schema.TryGetProperty("exclusiveMinimum", out var exclusiveMinimum) && value <= exclusiveMinimum.GetDecimal())
            diagnostics.Add(new("ppj.schema.minimum", $"Number must be greater than {exclusiveMinimum.GetRawText()}.", path.ToString()));
        if (schema.TryGetProperty("exclusiveMaximum", out var exclusiveMaximum) && value >= exclusiveMaximum.GetDecimal())
            diagnostics.Add(new("ppj.schema.maximum", $"Number must be less than {exclusiveMaximum.GetRawText()}.", path.ToString()));
        if (schema.TryGetProperty("multipleOf", out var multipleOf))
        {
            var factor = multipleOf.GetDecimal();
            if (factor != 0 && value % factor != 0)
                diagnostics.Add(new("ppj.schema.multiple", $"Number must be a multiple of {multipleOf.GetRawText()}.", path.ToString()));
        }
    }

    private static HashSet<string> CollectDeclaredProperties(JsonElement schema)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        CollectDeclaredProperties(schema, result, new HashSet<string>(StringComparer.Ordinal));
        return result;
    }

    private static void CollectDeclaredProperties(JsonElement schema, HashSet<string> result, HashSet<string> visitedRefs)
    {
        if (schema.ValueKind != JsonValueKind.Object) return;
        if (schema.TryGetProperty("$ref", out var reference))
        {
            var value = reference.GetString()!;
            if (visitedRefs.Add(value))
                CollectDeclaredProperties(ResolveReference(value), result, visitedRefs);
        }
        if (schema.TryGetProperty("properties", out var properties))
        {
            foreach (var property in properties.EnumerateObject())
                result.Add(property.Name);
        }
        if (schema.TryGetProperty("allOf", out var allOf))
        {
            foreach (var child in allOf.EnumerateArray())
                CollectDeclaredProperties(child, result, visitedRefs);
        }
    }

    private static int AppendProperty(StringBuilder path, string property)
    {
        var length = path.Length;
        if (property.All(character => char.IsAsciiLetterOrDigit(character) || character is '_' or '-'))
            path.Append('.').Append(property);
        else
            path.Append("['").Append(property.Replace("'", "\\'", StringComparison.Ordinal)).Append("']");
        return length;
    }

    private static int AppendIndex(StringBuilder path, int index)
    {
        var length = path.Length;
        path.Append('[').Append(index).Append(']');
        return length;
    }

    private static JsonElement ResolveReference(string reference)
    {
        const string prefix = "#/$defs/";
        if (!reference.StartsWith(prefix, StringComparison.Ordinal))
            throw new InvalidOperationException($"Unsupported PPJ schema reference {reference}.");
        var name = reference[prefix.Length..].Replace("~1", "/", StringComparison.Ordinal).Replace("~0", "~", StringComparison.Ordinal);
        if (!Definitions.Value.TryGetValue(name, out var result))
            throw new InvalidOperationException($"Missing PPJ schema definition {name}.");
        return result;
    }

    private static bool MatchesType(JsonElement instance, JsonElement type)
    {
        if (type.ValueKind == JsonValueKind.Array)
            return type.EnumerateArray().Any(candidate => MatchesType(instance, candidate));
        return type.GetString() switch
        {
            "object" => instance.ValueKind == JsonValueKind.Object,
            "array" => instance.ValueKind == JsonValueKind.Array,
            "string" => instance.ValueKind == JsonValueKind.String,
            "number" => instance.ValueKind == JsonValueKind.Number,
            "integer" => instance.ValueKind == JsonValueKind.Number && instance.TryGetInt64(out _),
            "boolean" => instance.ValueKind is JsonValueKind.True or JsonValueKind.False,
            "null" => instance.ValueKind == JsonValueKind.Null,
            _ => false,
        };
    }

    private static bool MatchesFormat(string value, string format) => format switch
    {
        "date-time" => DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _),
        "date" => DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _),
        "uri" => Uri.TryCreate(value, UriKind.Absolute, out var uri) && !string.IsNullOrEmpty(uri.Scheme),
        _ => true,
    };

    private static bool JsonEqual(JsonElement left, JsonElement right)
    {
        if (left.ValueKind != right.ValueKind) return false;
        return left.ValueKind switch
        {
            JsonValueKind.String => string.Equals(left.GetString(), right.GetString(), StringComparison.Ordinal),
            JsonValueKind.True or JsonValueKind.False or JsonValueKind.Null => true,
            _ => PpjCanonicalJson.Write(left).AsSpan().SequenceEqual(PpjCanonicalJson.Write(right)),
        };
    }

    private static string DescribeType(JsonElement type) =>
        type.ValueKind == JsonValueKind.Array
            ? string.Join(" or ", type.EnumerateArray().Select(item => item.GetString()))
            : type.GetString() ?? "declared type";

    private static string DescribeInstance(JsonElement instance) => instance.ValueKind switch
    {
        JsonValueKind.True or JsonValueKind.False => "boolean",
        JsonValueKind.Undefined => "undefined",
        _ => instance.ValueKind.ToString().ToLowerInvariant(),
    };
}

internal static class PpjCanonicalJson
{
    private sealed class PropertyNameComparer : IComparer<JsonProperty>
    {
        internal static readonly PropertyNameComparer Instance = new();
        public int Compare(JsonProperty left, JsonProperty right) =>
            StringComparer.Ordinal.Compare(left.Name, right.Name);
    }

    internal static byte[] Write(JsonElement value)
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Indented = false,
            SkipValidation = false,
        }))
        {
            WriteValue(writer, value);
        }
        return stream.ToArray();
    }

    private static void WriteValue(Utf8JsonWriter writer, JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                writer.WriteStartObject();
                var count = 0;
                foreach (var _ in value.EnumerateObject()) count++;
                if (count == 0)
                {
                    writer.WriteEndObject();
                    break;
                }
                var properties = ArrayPool<JsonProperty>.Shared.Rent(count);
                try
                {
                    var index = 0;
                    foreach (var property in value.EnumerateObject()) properties[index++] = property;
                    Array.Sort(properties, 0, count, PropertyNameComparer.Instance);
                    for (index = 0; index < count; index++)
                    {
                        var property = properties[index];
                        writer.WritePropertyName(property.Name);
                        WriteValue(writer, property.Value);
                    }
                    writer.WriteEndObject();
                }
                finally
                {
                    Array.Clear(properties, 0, count);
                    ArrayPool<JsonProperty>.Shared.Return(properties);
                }
                break;
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (var item in value.EnumerateArray())
                    WriteValue(writer, item);
                writer.WriteEndArray();
                break;
            case JsonValueKind.String:
                writer.WriteStringValue(value.GetString());
                break;
            case JsonValueKind.Number:
                if (value.TryGetInt64(out var signed)) writer.WriteNumberValue(signed);
                else if (value.TryGetUInt64(out var unsigned)) writer.WriteNumberValue(unsigned);
                else if (value.TryGetDecimal(out var decimalValue)) writer.WriteNumberValue(decimalValue);
                else writer.WriteNumberValue(value.GetDouble());
                break;
            case JsonValueKind.True:
                writer.WriteBooleanValue(true);
                break;
            case JsonValueKind.False:
                writer.WriteBooleanValue(false);
                break;
            case JsonValueKind.Null:
                writer.WriteNullValue();
                break;
            default:
                throw new InvalidOperationException("PPJ canonicalization received an undefined JSON value.");
        }
    }
}

internal static class PpjJsonPath
{
    internal static string Property(string path, string property) =>
        property.All(character => char.IsAsciiLetterOrDigit(character) || character is '_' or '-')
            ? $"{path}.{property}"
            : $"{path}['{property.Replace("'", "\\'", StringComparison.Ordinal)}']";
}
