using System.Globalization;
using System.Reflection;
using System.Text.Json;

namespace OfficeKit.Codec;

internal sealed record PpjIconPathCommand(char Operation, IReadOnlyList<double> Values);

internal sealed class PpjIconDefinition
{
    private readonly Lazy<IReadOnlyList<PpjIconPathCommand>> commands;

    internal PpjIconDefinition(double width, double height, string path)
    {
        Width = width;
        Height = height;
        commands = new(() => ParsePath(path), LazyThreadSafetyMode.ExecutionAndPublication);
    }

    internal double Width { get; }
    internal double Height { get; }
    internal IReadOnlyList<PpjIconPathCommand> Commands => commands.Value;

    private static IReadOnlyList<PpjIconPathCommand> ParsePath(string path)
    {
        var output = new List<PpjIconPathCommand>();
        var index = 0;
        while (index < path.Length)
        {
            var operation = path[index++];
            var parameterCount = operation switch
            {
                'M' or 'L' => 2,
                'C' => 6,
                'Z' => 0,
                _ => throw new InvalidOperationException($"Generated PPJ icon path contains unsupported command {operation}."),
            };
            var values = new double[parameterCount];
            for (var parameter = 0; parameter < parameterCount; parameter++)
            {
                while (index < path.Length && path[index] == ' ') index++;
                var start = index;
                while (index < path.Length && path[index] != ' ' && path[index] is not ('M' or 'L' or 'C' or 'Z')) index++;
                if (start == index || !double.TryParse(
                        path.AsSpan(start, index - start),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out values[parameter]) ||
                    !double.IsFinite(values[parameter]))
                    throw new InvalidOperationException("Generated PPJ icon path contains an invalid numeric coordinate.");
            }
            output.Add(new(operation, values));
            if (output.Count > 512)
                throw new InvalidOperationException("Generated PPJ icon path exceeds the 512-command compiler bound.");
        }
        if (output.Count == 0 || output[0].Operation != 'M')
            throw new InvalidOperationException("Generated PPJ icon path must begin with move-to.");
        return output;
    }
}

internal static class PpjIconCatalog
{
    private const string ResourceName = "OfficeKit.Ppj.FontAwesomeFreeIcons.json";
    private const string Schema = "office-kit/ppj-icon-catalog/v1";
    private const string Version = "7.3.1";
    private static readonly Lazy<IReadOnlyDictionary<string, PpjIconDefinition>> Definitions =
        new(Load, LazyThreadSafetyMode.ExecutionAndPublication);

    internal static bool Contains(string name) => Definitions.Value.ContainsKey(name);

    internal static PpjIconDefinition Resolve(string name)
    {
        if (Definitions.Value.TryGetValue(name, out var definition)) return definition;
        throw new CodecException(
            "ppj.icon.unknown",
            $"PPJ iconName {name} is not present in the pinned Font Awesome Free {Version} catalog.");
    }

    private static IReadOnlyDictionary<string, PpjIconDefinition> Load()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName) ??
            throw new InvalidOperationException($"Embedded PPJ icon catalog {ResourceName} is missing.");
        using var document = JsonDocument.Parse(stream);
        var root = document.RootElement;
        if (root.GetProperty("schema").GetString() != Schema ||
            root.GetProperty("source").GetProperty("version").GetString() != Version)
            throw new InvalidOperationException("Embedded PPJ icon catalog identity does not match the compiler contract.");

        var output = new Dictionary<string, PpjIconDefinition>(StringComparer.Ordinal);
        foreach (var property in root.GetProperty("icons").EnumerateObject())
        {
            var value = property.Value;
            var definition = new PpjIconDefinition(
                value.GetProperty("width").GetDouble(),
                value.GetProperty("height").GetDouble(),
                value.GetProperty("path").GetString()!);
            if (definition.Width <= 0 || definition.Height <= 0 || !output.TryAdd(property.Name, definition))
                throw new InvalidOperationException($"Embedded PPJ icon catalog has an invalid entry for {property.Name}.");
        }
        if (output.Count != 2_163)
            throw new InvalidOperationException($"Embedded PPJ icon catalog has {output.Count} entries; expected 2163.");
        return output;
    }
}
