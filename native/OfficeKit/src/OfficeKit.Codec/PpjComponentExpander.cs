using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeKit.Codec;

internal sealed record PpjExpandedNodeModel(
    string PageId,
    string Id,
    string SourceId,
    string Type,
    string? ComponentId,
    string? InstanceId,
    string? RepeatKey,
    string ProgramPath,
    int ZOrder);

internal sealed record PpjExpandedPageModel(
    string Id,
    IReadOnlyList<PpjElementModel> Elements,
    IReadOnlyList<JsonElement> ElementJson);

internal sealed record PpjExpansionResult(
    IReadOnlyList<PpjExpandedPageModel> Pages,
    IReadOnlyList<PpjExpandedNodeModel> Nodes,
    byte[] NodeMapJson,
    string NodeMapSha256,
    int ExpandedElementCount);

internal static class PpjComponentExpander
{
    internal static PpjExpansionResult? Expand(
        PpjProgramModel program,
        string programSha256,
        List<PpjDiagnostic> diagnostics)
    {
        var components = program.Components
            .GroupBy(item => item.Id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var pages = new List<PpjExpandedPageModel>(program.Pages.Count);
        var nodes = new List<PpjExpandedNodeModel>();
        var outputIds = new HashSet<string>(StringComparer.Ordinal);

        for (var pageIndex = 0; pageIndex < program.Pages.Count; pageIndex++)
        {
            var page = program.Pages[pageIndex];
            var expandedJson = new List<JsonElement>();
            for (var elementIndex = 0; elementIndex < page.Elements.Count; elementIndex++)
            {
                var path = $"$.pages[{pageIndex}].elements[{elementIndex}]";
                ExpandElement(
                    page.Elements[elementIndex],
                    path,
                    page.Id,
                    components,
                    Transform.Identity,
                    null,
                    null,
                    null,
                    1,
                    expandedJson,
                    nodes,
                    outputIds,
                    diagnostics);
                if (diagnostics.Count > 0 && nodes.Count > PpjProgramValidator.MaxExpandedElements)
                    return null;
            }
            var typed = new List<PpjElementModel>(expandedJson.Count);
            for (var index = 0; index < expandedJson.Count; index++)
            {
                try
                {
                    typed.Add(PpjProgramParser.ParseElement(expandedJson[index]));
                }
                catch (Exception exception) when (exception is InvalidOperationException or KeyNotFoundException or FormatException)
                {
                    diagnostics.Add(new(
                        "ppj.component.expandedModel",
                        "Expanded component output could not be projected into the typed element model.",
                        $"$.pages[{pageIndex}].elements[{index}]"));
                }
            }
            pages.Add(new(page.Id, typed, expandedJson));
        }

        if (diagnostics.Count > 0) return null;

        var nodeMapJson = WriteNodeMap(programSha256, nodes);
        var nodeMapSha256 = Convert.ToHexString(SHA256.HashData(nodeMapJson)).ToLowerInvariant();
        return new(pages, nodes, nodeMapJson, nodeMapSha256, nodes.Count);
    }

    private static void ExpandElement(
        PpjElementModel element,
        string path,
        string pageId,
        IReadOnlyDictionary<string, PpjComponentModel> components,
        Transform transform,
        string? identityPrefix,
        string? componentId,
        string? repeatKey,
        int depth,
        List<JsonElement> output,
        List<PpjExpandedNodeModel> nodes,
        HashSet<string> outputIds,
        List<PpjDiagnostic> diagnostics)
    {
        if (nodes.Count >= PpjProgramValidator.MaxExpandedElements)
        {
            diagnostics.Add(new(
                "ppj.component.elementBudget",
                $"Expanded element count exceeds {PpjProgramValidator.MaxExpandedElements}.",
                path));
            return;
        }
        if (depth > PpjProgramValidator.MaxComponentDepth)
        {
            diagnostics.Add(new(
                "ppj.component.depth",
                $"Component expansion exceeds depth {PpjProgramValidator.MaxComponentDepth}.",
                path));
            return;
        }

        if (element is PpjComponentElementModel instance)
        {
            ExpandComponent(
                instance,
                path,
                pageId,
                components,
                transform,
                identityPrefix,
                depth,
                output,
                nodes,
                outputIds,
                diagnostics);
            return;
        }
        if (element is PpjSlotElementModel)
        {
            diagnostics.Add(new("ppj.component.unfilledSlot", "An unfilled slot reached component output.", path));
            return;
        }

        var json = CloneObject(element.Raw);
        var sourceId = element.Id;
        var outputId = identityPrefix is null ? sourceId : DerivedId(identityPrefix, sourceId);
        if (identityPrefix is not null)
            RewriteLocalReferences(json, identityPrefix);
        json["id"] = outputId;
        WriteFrame(json, transform.Apply(element.Frame));

        if (element is PpjGroupElementModel group)
        {
            var children = new JsonArray();
            for (var index = 0; index < group.Elements.Count; index++)
            {
                var childOutput = new List<JsonElement>();
                ExpandElement(
                    group.Elements[index],
                    $"{path}.elements[{index}]",
                    pageId,
                    components,
                    transform,
                    identityPrefix,
                    componentId,
                    repeatKey,
                    depth,
                    childOutput,
                    nodes,
                    outputIds,
                    diagnostics);
                foreach (var child in childOutput)
                    children.Add(JsonNode.Parse(child.GetRawText()));
            }
            json["elements"] = children;
        }

        if (!outputIds.Add(outputId))
        {
            diagnostics.Add(new("ppj.component.expandedId", $"Expanded element ID {outputId} collides with another output element.", $"{path}.id"));
            return;
        }

        var zOrder = output.Count;
        var serialized = ToElement(json);
        output.Add(serialized);
        nodes.Add(new(pageId, outputId, sourceId, element.Type, componentId, identityPrefix, repeatKey, path, zOrder));
    }

    private static void ExpandComponent(
        PpjComponentElementModel instance,
        string path,
        string pageId,
        IReadOnlyDictionary<string, PpjComponentModel> components,
        Transform parentTransform,
        string? parentIdentity,
        int depth,
        List<JsonElement> output,
        List<PpjExpandedNodeModel> nodes,
        HashSet<string> outputIds,
        List<PpjDiagnostic> diagnostics)
    {
        if (!components.TryGetValue(instance.ComponentId, out var definition)) return;
        var instanceId = parentIdentity is null ? instance.Id : DerivedId(parentIdentity, instance.Id);
        var repetitions = instance.Repeat?.Items.Count ?? 1;
        for (var repeatIndex = 0; repeatIndex < repetitions; repeatIndex++)
        {
            var item = instance.Repeat is null ? null : instance.Repeat.Items[repeatIndex];
            var arguments = MergeArguments(definition, instance.Arguments, item?.Arguments);
            if (instance.When is not null && !Evaluate(instance.When, arguments)) continue;
            var key = item?.Key;
            var identity = DerivedId(instanceId, instance.ComponentId, key ?? "single");
            var itemTransform = ComponentTransform(definition.Frame, instance.Frame, instance.Repeat, repeatIndex)
                .Then(parentTransform);

            for (var elementIndex = 0; elementIndex < definition.Elements.Count; elementIndex++)
            {
                var templateElement = definition.Elements[elementIndex];
                var elementPath = $"{path}#component({definition.Id}).elements[{elementIndex}]";
                if (templateElement is PpjSlotElementModel slot)
                {
                    if (!instance.Slots.TryGetValue(slot.SlotId, out var supplied)) continue;
                    for (var slotIndex = 0; slotIndex < supplied.Count; slotIndex++)
                    {
                        ExpandElement(
                            supplied[slotIndex],
                            $"{path}.slots.{slot.SlotId}[{slotIndex}]",
                            pageId,
                            components,
                            itemTransform,
                            identity,
                            definition.Id,
                            key,
                            depth + 1,
                            output,
                            nodes,
                            outputIds,
                            diagnostics);
                    }
                    continue;
                }

                var bound = BindElement(templateElement.Raw, definition, arguments, instance.VariantId, instance.Slots, elementPath, diagnostics);
                if (bound is null) continue;
                PpjElementModel typed;
                try
                {
                    typed = PpjProgramParser.ParseElement(bound.Value);
                }
                catch (Exception exception) when (exception is InvalidOperationException or KeyNotFoundException or FormatException)
                {
                    diagnostics.Add(new("ppj.component.bindingOutput", "A component binding produced an invalid typed element.", elementPath));
                    continue;
                }
                ExpandElement(
                    typed,
                    elementPath,
                    pageId,
                    components,
                    itemTransform,
                    identity,
                    definition.Id,
                    key,
                    depth + 1,
                    output,
                    nodes,
                    outputIds,
                    diagnostics);
            }
        }
    }

    private static JsonElement? BindElement(
        JsonElement raw,
        PpjComponentModel definition,
        IReadOnlyDictionary<string, JsonElement> arguments,
        string? requestedVariant,
        IReadOnlyDictionary<string, IReadOnlyList<PpjElementModel>> slots,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var element = CloneObject(raw);
        ReplaceNestedSlots(element, slots);
        foreach (var binding in definition.Bindings)
        {
            var target = FindById(element, binding.TargetId);
            if (target is not null && arguments.TryGetValue(binding.Parameter, out var value))
                ApplyValue(target, binding.Field, value, path, diagnostics);
        }

        var variants = requestedVariant is null
            ? definition.Variants
            : definition.Variants.Where(item => item.Id == requestedVariant).ToArray();
        foreach (var variant in variants)
        {
            if (!Evaluate(variant.When, arguments)) continue;
            foreach (var assignment in variant.Assignments)
            {
                var target = FindById(element, assignment.TargetId);
                if (target is not null)
                    ApplyValue(target, assignment.Field, assignment.Value, path, diagnostics);
            }
        }

        return ToElement(element);
    }

    private static IReadOnlyDictionary<string, JsonElement> MergeArguments(
        PpjComponentModel definition,
        IReadOnlyDictionary<string, JsonElement> instance,
        IReadOnlyDictionary<string, JsonElement>? repeat)
    {
        var result = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var parameter in definition.Parameters)
            if (parameter.Default is { } defaultValue) result[parameter.Name] = defaultValue;
        foreach (var value in instance) result[value.Key] = value.Value;
        if (repeat is not null)
            foreach (var value in repeat) result[value.Key] = value.Value;
        return result;
    }

    private static bool Evaluate(PpjConditionModel condition, IReadOnlyDictionary<string, JsonElement> arguments)
    {
        var present = arguments.TryGetValue(condition.Argument, out var actual) && actual.ValueKind != JsonValueKind.Null;
        return condition.Operator switch
        {
            "present" => present,
            "absent" => !present,
            "equals" => present && condition.Value is { } expected && JsonEqual(actual, expected),
            "notEquals" => !present || condition.Value is not { } expected || !JsonEqual(actual, expected),
            _ => false,
        };
    }

    private static void ApplyValue(JsonObject target, string field, JsonElement value, string path, List<PpjDiagnostic> diagnostics)
    {
        try
        {
            var node = JsonNode.Parse(value.GetRawText());
            switch (field)
            {
                case "text":
                    target["text"] = Unwrap(node, "text");
                    break;
                case "fill":
                    var fillOwner = target["type"]?.GetValue<string>() == "shape"
                        ? EnsureObject(target, "style")
                        : target;
                    fillOwner["fill"] = ToFill(node);
                    break;
                case "stroke":
                    var strokeOwner = target["type"]?.GetValue<string>() == "shape"
                        ? EnsureObject(target, "style")
                        : target;
                    strokeOwner["stroke"] = ToStroke(node);
                    break;
                case "opacity":
                    var opacityOwner = target["type"]?.GetValue<string>() == "shape"
                        ? EnsureObject(target, "style")
                        : target;
                    opacityOwner["opacity"] = node;
                    break;
                case "frame.x": SetFrame(target, "x", node); break;
                case "frame.y": SetFrame(target, "y", node); break;
                case "frame.width": SetFrame(target, "width", node); break;
                case "frame.height": SetFrame(target, "height", node); break;
                case "image.asset": target["asset"] = Unwrap(node, "asset"); break;
                case "chart.title": target["title"] = Unwrap(node, "text"); break;
                case "accessibility.description": EnsureObject(target, "accessibility")["description"] = node; break;
                default:
                    diagnostics.Add(new("ppj.component.bindingUnsupported", $"Binding field {field} is not implemented by PPJ v1 expansion.", path));
                    break;
            }
        }
        catch (Exception exception) when (exception is InvalidOperationException or FormatException)
        {
            diagnostics.Add(new("ppj.component.bindingValue", $"Binding field {field} received an incompatible value.", path));
        }
    }

    private static JsonNode? Unwrap(JsonNode? node, string property) =>
        node is JsonObject wrapper && wrapper.TryGetPropertyValue(property, out var value)
            ? value?.DeepClone()
            : node?.DeepClone();

    private static JsonNode? ToFill(JsonNode? node)
    {
        if (node is JsonObject wrapper && wrapper.TryGetPropertyValue("color", out var color))
            return new JsonObject { ["type"] = "solid", ["color"] = color?.DeepClone() };
        return node?.DeepClone();
    }

    private static JsonNode? ToStroke(JsonNode? node)
    {
        if (node is JsonObject wrapper && wrapper.TryGetPropertyValue("color", out var color))
            return new JsonObject { ["color"] = color?.DeepClone(), ["width"] = 1.0, ["dash"] = "solid" };
        return node?.DeepClone();
    }

    private static void SetFrame(JsonObject target, string name, JsonNode? value) =>
        EnsureObject(target, "frame")[name] = value?.DeepClone();

    private static JsonObject EnsureObject(JsonObject owner, string name)
    {
        if (owner[name] is JsonObject existing) return existing;
        var created = new JsonObject();
        owner[name] = created;
        return created;
    }

    private static JsonObject? FindById(JsonNode? node, string id)
    {
        if (node is JsonObject value)
        {
            if (value["id"] is JsonValue candidate && candidate.TryGetValue<string>(out var text) && text == id)
                return value;
            foreach (var child in value)
            {
                var found = FindById(child.Value, id);
                if (found is not null) return found;
            }
        }
        else if (node is JsonArray array)
        {
            foreach (var child in array)
            {
                var found = FindById(child, id);
                if (found is not null) return found;
            }
        }
        return null;
    }

    private static void ReplaceNestedSlots(
        JsonNode? node,
        IReadOnlyDictionary<string, IReadOnlyList<PpjElementModel>> slots)
    {
        if (node is JsonObject value)
        {
            foreach (var property in value.ToArray())
                ReplaceNestedSlots(property.Value, slots);
        }
        else if (node is JsonArray array)
        {
            for (var index = array.Count - 1; index >= 0; index--)
            {
                if (array[index] is JsonObject child && child["type"]?.GetValue<string>() == "slot")
                {
                    var slotId = child["slot"]?.GetValue<string>();
                    array.RemoveAt(index);
                    if (slotId is not null && slots.TryGetValue(slotId, out var supplied))
                    {
                        for (var suppliedIndex = supplied.Count - 1; suppliedIndex >= 0; suppliedIndex--)
                            array.Insert(index, JsonNode.Parse(supplied[suppliedIndex].Raw.GetRawText()));
                    }
                }
                else
                {
                    ReplaceNestedSlots(array[index], slots);
                }
            }
        }
    }

    private static Transform ComponentTransform(
        PpjFrameModel definition,
        PpjFrameModel instance,
        PpjRepeatModel? repeat,
        int repeatIndex)
    {
        if (repeat is null)
        {
            return new(
                instance.X - definition.X * (instance.Width / definition.Width),
                instance.Y - definition.Y * (instance.Height / definition.Height),
                instance.Width / definition.Width,
                instance.Height / definition.Height);
        }

        var direction = repeat.Direction ?? "vertical";
        if (direction == "horizontal")
        {
            var scale = instance.Height / definition.Height;
            var itemWidth = definition.Width * scale;
            return new(
                instance.X + repeatIndex * (itemWidth + repeat.Gap) - definition.X * scale,
                instance.Y - definition.Y * scale,
                scale,
                scale);
        }
        else
        {
            var scale = instance.Width / definition.Width;
            var itemHeight = definition.Height * scale;
            return new(
                instance.X - definition.X * scale,
                instance.Y + repeatIndex * (itemHeight + repeat.Gap) - definition.Y * scale,
                scale,
                scale);
        }
    }

    private static void RewriteLocalReferences(JsonObject element, string identityPrefix)
        => Rewrite(element, identityPrefix);

    private static void Rewrite(JsonNode? node, string identityPrefix)
    {
        if (node is JsonObject value)
        {
            foreach (var property in value.ToArray())
            {
                if (property.Value is JsonValue scalar && scalar.TryGetValue<string>(out var text) &&
                    property.Key is "id" or "parent" or "element" or "target" or "from" or "to")
                {
                    value[property.Key] = DerivedId(identityPrefix, text);
                }
                else
                {
                    Rewrite(property.Value, identityPrefix);
                }
            }
        }
        else if (node is JsonArray array)
        {
            foreach (var child in array) Rewrite(child, identityPrefix);
        }
    }

    private static string DerivedId(params string[] parts)
    {
        var value = string.Join("::", parts.Where(part => !string.IsNullOrEmpty(part)));
        if (value.Length <= 128) return value;
        var suffix = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant()[..20];
        return $"{value[..Math.Min(105, value.Length)]}::{suffix}";
    }

    private static JsonObject CloneObject(JsonElement element) =>
        JsonNode.Parse(element.GetRawText())?.AsObject()
        ?? throw new InvalidOperationException("PPJ element must be an object.");

    private static JsonElement ToElement(JsonObject value)
    {
        using var document = JsonDocument.Parse(value.ToJsonString());
        return document.RootElement.Clone();
    }

    private static void WriteFrame(JsonObject owner, PpjFrameModel frame)
    {
        var value = EnsureObject(owner, "frame");
        value["x"] = frame.X;
        value["y"] = frame.Y;
        value["width"] = frame.Width;
        value["height"] = frame.Height;
        if (frame.Rotation != 0) value["rotation"] = frame.Rotation; else value.Remove("rotation");
        if (frame.FlipH) value["flipH"] = true; else value.Remove("flipH");
        if (frame.FlipV) value["flipV"] = true; else value.Remove("flipV");
    }

    private static byte[] WriteNodeMap(string programSha256, IReadOnlyList<PpjExpandedNodeModel> nodes)
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false }))
        {
            writer.WriteStartObject();
            writer.WriteString("schema", "office-kit/ppj-node-map/v1");
            writer.WriteString("programSha256", programSha256);
            writer.WriteNumber("expandedElementCount", nodes.Count);
            writer.WriteStartArray("nodes");
            foreach (var node in nodes)
            {
                writer.WriteStartObject();
                writer.WriteString("page", node.PageId);
                writer.WriteString("id", node.Id);
                writer.WriteString("sourceId", node.SourceId);
                writer.WriteString("type", node.Type);
                if (node.ComponentId is not null) writer.WriteString("component", node.ComponentId);
                if (node.InstanceId is not null) writer.WriteString("instance", node.InstanceId);
                if (node.RepeatKey is not null) writer.WriteString("repeatKey", node.RepeatKey);
                writer.WriteString("path", node.ProgramPath);
                writer.WriteNumber("zOrder", node.ZOrder);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return stream.ToArray();
    }

    private static bool JsonEqual(JsonElement left, JsonElement right) =>
        PpjCanonicalJson.Write(left).AsSpan().SequenceEqual(PpjCanonicalJson.Write(right));

    private readonly record struct Transform(double OffsetX, double OffsetY, double ScaleX, double ScaleY)
    {
        internal static Transform Identity => new(0, 0, 1, 1);

        internal PpjFrameModel Apply(PpjFrameModel frame) => new(
            OffsetX + frame.X * ScaleX,
            OffsetY + frame.Y * ScaleY,
            frame.Width * ScaleX,
            frame.Height * ScaleY,
            frame.Rotation,
            frame.FlipH,
            frame.FlipV);

        internal Transform Then(Transform outer) => new(
            outer.OffsetX + OffsetX * outer.ScaleX,
            outer.OffsetY + OffsetY * outer.ScaleY,
            ScaleX * outer.ScaleX,
            ScaleY * outer.ScaleY);
    }
}
