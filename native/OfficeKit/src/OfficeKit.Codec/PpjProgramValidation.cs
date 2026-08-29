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
    PpjExpansionResult? Expansion)
{
    internal bool IsValid => Program is not null && Diagnostics.Count == 0;
}

internal static class PpjProgramValidator
{
    internal const int MaxSourceBytes = 16 * 1024 * 1024;
    internal const int MaxPages = 512;
    internal const int MaxExpandedElements = 100_000;
    internal const int MaxRepeatItems = 1_024;
    internal const int MaxComponentDepth = 16;
    private const int MaxJsonDepth = 96;

    internal static PpjValidationResult Validate(ReadOnlySpan<byte> json)
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
        if (json.Length >= 3 && json[0] == 0xef && json[1] == 0xbb && json[2] == 0xbf)
        {
            diagnostics.Add(new("ppj.utf8Bom", "PPJ must be UTF-8 JSON without a byte-order mark.", "$"));
            return Invalid(diagnostics);
        }

        JsonElement root;
        try
        {
            using var document = JsonDocument.Parse(json.ToArray(), new JsonDocumentOptions
            {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = MaxJsonDepth,
            });
            root = document.RootElement.Clone();
        }
        catch (JsonException exception)
        {
            diagnostics.Add(new(
                "ppj.invalidJson",
                $"PPJ is not strict UTF-8 JSON: {exception.Message}",
                exception.Path ?? "$"));
            return Invalid(diagnostics);
        }

        FindDuplicateProperties(root, "$", diagnostics);
        if (diagnostics.Count == 0)
            PpjJsonSchemaValidator.Validate(root, diagnostics);
        if (diagnostics.Count != 0)
            return Invalid(diagnostics);

        PpjProgramModel program;
        try
        {
            program = PpjProgramParser.Parse(root);
        }
        catch (Exception exception) when (exception is KeyNotFoundException or InvalidOperationException or FormatException)
        {
            diagnostics.Add(new("ppj.modelProjection", "Validated PPJ could not be projected into the native typed model.", "$"));
            return Invalid(diagnostics);
        }

        PpjSemanticValidator.Validate(program, diagnostics);
        if (diagnostics.Count != 0)
            return Invalid(diagnostics);

        var canonical = PpjCanonicalJson.Write(root);
        var hash = Convert.ToHexString(SHA256.HashData(canonical)).ToLowerInvariant();
        var expansion = PpjComponentExpander.Expand(program, hash, diagnostics);
        if (diagnostics.Count != 0 || expansion is null)
            return Invalid(diagnostics);
        return new PpjValidationResult(program, diagnostics, canonical, hash, expansion);
    }

    private static PpjValidationResult Invalid(IReadOnlyList<PpjDiagnostic> diagnostics) =>
        new(null, diagnostics, [], string.Empty, null);

    private static void FindDuplicateProperties(JsonElement value, string path, List<PpjDiagnostic> diagnostics)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in value.EnumerateObject())
            {
                var propertyPath = PpjJsonPath.Property(path, property.Name);
                if (!names.Add(property.Name))
                    diagnostics.Add(new("ppj.duplicateProperty", $"Property {property.Name} appears more than once.", propertyPath));
                FindDuplicateProperties(property.Value, propertyPath, diagnostics);
            }
        }
        else if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
                FindDuplicateProperties(item, $"{path}[{index++}]", diagnostics);
        }
    }
}

internal static class PpjJsonSchemaValidator
{
    private static readonly Lazy<JsonElement> Schema = new(LoadSchema, LazyThreadSafetyMode.ExecutionAndPublication);
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(100);

    internal static void Validate(JsonElement instance, List<PpjDiagnostic> diagnostics) =>
        ValidateNode(instance, Schema.Value, "$", diagnostics);

    private static JsonElement LoadSchema()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OfficeKit.Ppj.V1.Schema.json")
            ?? throw new InvalidOperationException("Embedded PPJ v1 schema is missing.");
        using var document = JsonDocument.Parse(stream);
        return document.RootElement.Clone();
    }

    private static void ValidateNode(JsonElement instance, JsonElement schema, string path, List<PpjDiagnostic> diagnostics)
    {
        if (schema.ValueKind == JsonValueKind.True) return;
        if (schema.ValueKind == JsonValueKind.False)
        {
            diagnostics.Add(new("ppj.schema.false", "Value is prohibited by the PPJ schema.", path));
            return;
        }
        if (schema.ValueKind != JsonValueKind.Object)
        {
            diagnostics.Add(new("ppj.schema.internal", "PPJ schema contains an invalid rule.", path));
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
                path));
            return;
        }

        if (schema.TryGetProperty("const", out var constant) && !JsonEqual(instance, constant))
            diagnostics.Add(new("ppj.schema.const", $"Value must equal {constant.GetRawText()}.", path));

        if (schema.TryGetProperty("enum", out var enumeration) && !enumeration.EnumerateArray().Any(item => JsonEqual(instance, item)))
            diagnostics.Add(new("ppj.schema.enum", "Value is not one of the allowed PPJ values.", path));

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
        string path,
        List<PpjDiagnostic> diagnostics,
        bool requireExactlyOne)
    {
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
            path));
    }

    private static void ValidateObject(JsonElement instance, JsonElement schema, string path, List<PpjDiagnostic> diagnostics)
    {
        var properties = instance.EnumerateObject().ToArray();
        if (schema.TryGetProperty("maxProperties", out var maximum) && properties.Length > maximum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"Object has {properties.Length} properties; maximum is {maximum.GetInt32()}.", path));

        if (schema.TryGetProperty("required", out var required))
        {
            foreach (var name in required.EnumerateArray().Select(item => item.GetString()!))
            {
                if (!instance.TryGetProperty(name, out _))
                    diagnostics.Add(new("ppj.schema.required", $"Required property {name} is missing.", PpjJsonPath.Property(path, name)));
            }
        }

        var declared = schema.TryGetProperty("properties", out var propertySchemas)
            ? propertySchemas.EnumerateObject().ToDictionary(item => item.Name, item => item.Value, StringComparer.Ordinal)
            : new Dictionary<string, JsonElement>(StringComparer.Ordinal);

        foreach (var property in properties)
        {
            var propertyPath = PpjJsonPath.Property(path, property.Name);
            if (declared.TryGetValue(property.Name, out var propertySchema))
                ValidateNode(property.Value, propertySchema, propertyPath, diagnostics);
            if (schema.TryGetProperty("propertyNames", out var propertyNameSchema))
                ValidateNode(JsonSerializer.SerializeToElement(property.Name), propertyNameSchema, propertyPath, diagnostics);
        }

        if (schema.TryGetProperty("additionalProperties", out var additional))
        {
            foreach (var property in properties.Where(item => !declared.ContainsKey(item.Name)))
            {
                var propertyPath = PpjJsonPath.Property(path, property.Name);
                if (additional.ValueKind == JsonValueKind.False)
                    diagnostics.Add(new("ppj.schema.unknownField", $"Unknown property {property.Name} is not allowed.", propertyPath));
                else if (additional.ValueKind == JsonValueKind.Object || additional.ValueKind is JsonValueKind.True or JsonValueKind.False)
                    ValidateNode(property.Value, additional, propertyPath, diagnostics);
            }
        }

        if (schema.TryGetProperty("unevaluatedProperties", out var unevaluated) && unevaluated.ValueKind == JsonValueKind.False)
        {
            var evaluated = CollectDeclaredProperties(schema);
            foreach (var property in properties.Where(item => !evaluated.Contains(item.Name)))
                diagnostics.Add(new("ppj.schema.unknownField", $"Unknown property {property.Name} is not allowed.", PpjJsonPath.Property(path, property.Name)));
        }
    }

    private static void ValidateArray(JsonElement instance, JsonElement schema, string path, List<PpjDiagnostic> diagnostics)
    {
        var items = instance.EnumerateArray().ToArray();
        if (schema.TryGetProperty("minItems", out var minimum) && items.Length < minimum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"Array has {items.Length} items; minimum is {minimum.GetInt32()}.", path));
        if (schema.TryGetProperty("maxItems", out var maximum) && items.Length > maximum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"Array has {items.Length} items; maximum is {maximum.GetInt32()}.", path));
        if (schema.TryGetProperty("uniqueItems", out var unique) && unique.GetBoolean())
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 0; index < items.Length; index++)
            {
                var key = Convert.ToBase64String(PpjCanonicalJson.Write(items[index]));
                if (!seen.Add(key))
                    diagnostics.Add(new("ppj.schema.unique", "Array item duplicates an earlier value.", $"{path}[{index}]"));
            }
        }
        if (schema.TryGetProperty("items", out var itemSchema))
        {
            for (var index = 0; index < items.Length; index++)
                ValidateNode(items[index], itemSchema, $"{path}[{index}]", diagnostics);
        }
    }

    private static void ValidateString(string value, JsonElement schema, string path, List<PpjDiagnostic> diagnostics)
    {
        var length = value.EnumerateRunes().Count();
        if (schema.TryGetProperty("minLength", out var minimum) && length < minimum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"String has {length} characters; minimum is {minimum.GetInt32()}.", path));
        if (schema.TryGetProperty("maxLength", out var maximum) && length > maximum.GetInt32())
            diagnostics.Add(new("ppj.schema.limit", $"String has {length} characters; maximum is {maximum.GetInt32()}.", path));
        if (schema.TryGetProperty("pattern", out var pattern))
        {
            try
            {
                if (!Regex.IsMatch(value, pattern.GetString()!, RegexOptions.CultureInvariant, RegexTimeout))
                    diagnostics.Add(new("ppj.schema.pattern", "String does not match the required PPJ pattern.", path));
            }
            catch (RegexMatchTimeoutException)
            {
                diagnostics.Add(new("ppj.schema.pattern", "String pattern validation exceeded its bounded time budget.", path));
            }
        }
        if (schema.TryGetProperty("format", out var format) && !MatchesFormat(value, format.GetString()!))
            diagnostics.Add(new("ppj.schema.format", $"String is not a valid {format.GetString()} value.", path));
    }

    private static void ValidateNumber(JsonElement instance, JsonElement schema, string path, List<PpjDiagnostic> diagnostics)
    {
        if (!instance.TryGetDecimal(out var value))
        {
            var floating = instance.GetDouble();
            if (!double.IsFinite(floating))
                diagnostics.Add(new("ppj.schema.number", "Number must be finite.", path));
            return;
        }
        if (schema.TryGetProperty("minimum", out var minimum) && value < minimum.GetDecimal())
            diagnostics.Add(new("ppj.schema.minimum", $"Number must be at least {minimum.GetRawText()}.", path));
        if (schema.TryGetProperty("maximum", out var maximum) && value > maximum.GetDecimal())
            diagnostics.Add(new("ppj.schema.maximum", $"Number must be at most {maximum.GetRawText()}.", path));
        if (schema.TryGetProperty("exclusiveMinimum", out var exclusiveMinimum) && value <= exclusiveMinimum.GetDecimal())
            diagnostics.Add(new("ppj.schema.minimum", $"Number must be greater than {exclusiveMinimum.GetRawText()}.", path));
        if (schema.TryGetProperty("exclusiveMaximum", out var exclusiveMaximum) && value >= exclusiveMaximum.GetDecimal())
            diagnostics.Add(new("ppj.schema.maximum", $"Number must be less than {exclusiveMaximum.GetRawText()}.", path));
        if (schema.TryGetProperty("multipleOf", out var multipleOf))
        {
            var factor = multipleOf.GetDecimal();
            if (factor != 0 && value % factor != 0)
                diagnostics.Add(new("ppj.schema.multiple", $"Number must be a multiple of {multipleOf.GetRawText()}.", path));
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

    private static JsonElement ResolveReference(string reference)
    {
        const string prefix = "#/$defs/";
        if (!reference.StartsWith(prefix, StringComparison.Ordinal))
            throw new InvalidOperationException($"Unsupported PPJ schema reference {reference}.");
        var name = reference[prefix.Length..].Replace("~1", "/", StringComparison.Ordinal).Replace("~0", "~", StringComparison.Ordinal);
        if (!Schema.Value.GetProperty("$defs").TryGetProperty(name, out var result))
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

    private static bool JsonEqual(JsonElement left, JsonElement right) =>
        PpjCanonicalJson.Write(left).AsSpan().SequenceEqual(PpjCanonicalJson.Write(right));

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
                foreach (var property in value.EnumerateObject().OrderBy(item => item.Name, StringComparer.Ordinal))
                {
                    writer.WritePropertyName(property.Name);
                    WriteValue(writer, property.Value);
                }
                writer.WriteEndObject();
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
