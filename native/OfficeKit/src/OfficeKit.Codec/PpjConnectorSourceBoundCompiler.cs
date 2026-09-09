using System.Globalization;
using Google.Protobuf;
using Google.Protobuf.Reflection;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static partial class PpjSourceBoundPresentationCompiler
{
    private sealed record ConnectorSourceBinding(PpjElementModel Program, PresentationElement Wire, string? Parent);

    private static Dictionary<string, ConnectorSourceBinding> IndexConnectorSourceBindings(
        IReadOnlyList<PpjElementModel> models, IEnumerable<PresentationElement> native)
    {
        var result = new Dictionary<string, ConnectorSourceBinding>(StringComparer.Ordinal);
        void Add(IReadOnlyList<PpjElementModel> elements, IEnumerable<PresentationElement> wire, string? parent)
        {
            var items = wire.ToArray();
            if (elements.Count != items.Length) throw Unsupported("$.pages", "inconsistent connector target topology");
            for (var i = 0; i < elements.Count; i++)
            {
                result.Add(elements[i].Id, new(elements[i], items[i], parent));
                if (elements[i] is PpjGroupElementModel group && items[i].Group is { } nativeGroup)
                    Add(group.Elements, nativeGroup.Children, group.Id);
            }
        }
        Add(models, native, null);
        return result;
    }

    private static bool ConnectorEndpointsChanged(PpjConnectorElementModel before, PpjConnectorElementModel after) =>
        PropertyChanged(before.Raw, after.Raw, "from") || PropertyChanged(before.Raw, after.Raw, "to");

    private static bool ApplyConnectorDependencies(PpjPageModel before, PpjPageModel after,
        IReadOnlyDictionary<string, ConnectorSourceBinding> bindings, MutationState mutations,
        ISet<string> changedNodes, string path)
    {
        var models = new Dictionary<string, PpjElementModel>(StringComparer.Ordinal);
        void Index(IEnumerable<PpjElementModel> elements)
        {
            foreach (var element in elements)
            {
                models.Add(element.Id, element);
                if (element is PpjGroupElementModel group) Index(group.Elements);
            }
        }
        Index(after.Elements);
        var connectors = models.Values.OfType<PpjConnectorElementModel>().Where(model =>
            bindings.TryGetValue(model.Id, out var source) && source.Program is PpjConnectorElementModel old &&
            (source.Wire.Connector?.StartFrameAnchor is not null || source.Wire.Connector?.EndFrameAnchor is not null ||
                model.From.ConnectionSite is not null || model.To.ConnectionSite is not null ||
                ConnectorEndpointsChanged(old, model))).ToArray();
        if (connectors.Length == 0) return false;

        var dependencies = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (var connector in connectors)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            void Add(string? id)
            {
                while (id is not null && ids.Add(id) && bindings.TryGetValue(id, out var bound)) id = bound.Parent;
            }
            Add(connector.Id);
            Add(connector.From.ElementId);
            Add(connector.To.ElementId);
            dependencies.Add(connector.Id, ids);
        }
        var affectedIds = dependencies.Values.SelectMany(ids => ids).ToHashSet(StringComparer.Ordinal);
        var siteTargets = connectors.SelectMany(c => new[] { c.From, c.To })
            .Where(e => e.ConnectionSite is not null).Select(e => e.ElementId!).ToHashSet(StringComparer.Ordinal);
        var frames = new Dictionary<string, PpjFrameModel>(StringComparer.Ordinal);
        var childFrames = new Dictionary<string, PpjFrameModel>(StringComparer.Ordinal);
        var leafChangedIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < mutations.NativeLeaves.Count; index++)
        {
            var mutation = mutations.NativeLeaves[index];
            if (siteTargets.Contains(mutation.ProgramElementId) && mutation.LeafKind.StartsWith("customGeometry", StringComparison.Ordinal))
                throw Unsupported(path, "bound custom-site geometry native leaves cannot participate in endpoint recomputation; edit semantic geometry instead");
            if (mutation.FrameFastPath || !affectedIds.Contains(mutation.ProgramElementId) ||
                !IsConnectorFrameLeaf(mutation.LeafKind)) continue;
            var model = models[mutation.ProgramElementId];
            if (model is PpjConnectorElementModel)
                throw Unsupported(path, "an attached connector frame is derived; edit endpoints instead of frame leaves");
            if (bindings[model.Id].Wire.Source?.Editable != true)
                throw Unsupported(path, "a connector dependency frame is not safe for semantic export");
            var child = mutation.LeafKind.StartsWith("child", StringComparison.Ordinal);
            var frame = child ? childFrames.GetValueOrDefault(model.Id, ((PpjGroupElementModel)model).ChildFrame)
                : frames.GetValueOrDefault(model.Id, model.Frame);
            var next = ConnectorLeafFrame(frame, mutation);
            ApplyConnectorFrameLeafToWire(mutation, path);
            if (child) childFrames[model.Id] = next; else frames[model.Id] = next;
            // The exact leaf was already re-proved against the fresh source.
            // Its complete frame state is now also represented in the native
            // candidate, so dependent connectors can share ordinary export.
            mutations.NativeLeaves[index] = mutation with { FrameFastPath = true };
            leafChangedIds.Add(model.Id);
        }

        PpjConnectorEndpointResolver? resolver = null;
        var changed = false;
        foreach (var connector in connectors)
        {
            var source = bindings[connector.Id];
            var old = (PpjConnectorElementModel)source.Program;
            var endpointEdit = ConnectorEndpointsChanged(old, connector);
            var dependencyEdit = dependencies[connector.Id].Any(id => leafChangedIds.Contains(id) ||
                models.TryGetValue(id, out var model) && bindings.TryGetValue(id, out var original) &&
                (FrameChanged(original.Program, model) || siteTargets.Contains(id) && PropertyChanged(original.Program.Raw, model.Raw, "geometry") || original.Program is PpjGroupElementModel oldGroup &&
                    model is PpjGroupElementModel group && ChildFrameChanged(oldGroup, group)));
            if (!endpointEdit && !dependencyEdit) continue; // Byte-exact no-op, including native rounding.
            RequireCapability(connector, "setConnectorEndpoints", path + "." + connector.Id);
            var target = source.Wire.Connector;
            if (target.StartTargetId.Length != 0 && old.From.ConnectionSite is null ||
                target.EndTargetId.Length != 0 && old.To.ConnectionSite is null)
                throw Unsupported(path, "native geometry connection sites remain source-owned");
            resolver ??= new PpjConnectorEndpointResolver(after.Elements, frames, childFrames);
            var resolved = resolver.Resolve(connector);
            var candidate = target.Clone();
            candidate.StartXEmu = resolved.StartX;
            candidate.StartYEmu = resolved.StartY;
            candidate.EndXEmu = resolved.EndX;
            candidate.EndYEmu = resolved.EndY;
            PresentationConnectorFrameAnchor? Anchor(PpjConnectorEndpointModel endpoint)
            {
                if (endpoint.ConnectionSite is not null || endpoint.ElementId is not { } id) return null;
                if (!models.ContainsKey(id) || !bindings.TryGetValue(id, out var bound))
                    throw Unsupported(path, $"connector target {id} has no retained native identity");
                return new() { TargetId = bound.Wire.Id, Anchor = endpoint.Anchor! };
            }
            candidate.StartFrameAnchor = Anchor(connector.From);
            candidate.EndFrameAnchor = Anchor(connector.To);
            string SiteTarget(PpjConnectorEndpointModel endpoint)
            {
                if (endpoint.ConnectionSite is null) return string.Empty;
                if (endpoint.ElementId is not { } id || !models.ContainsKey(id) || !bindings.TryGetValue(id, out var bound))
                    throw Unsupported(path, "custom-site target has no retained native identity");
                return bound.Wire.Id;
            }
            candidate.StartTargetId = SiteTarget(connector.From);
            candidate.StartConnectionSiteIndex = connector.From.ConnectionSite ?? 0;
            candidate.EndTargetId = SiteTarget(connector.To);
            candidate.EndConnectionSiteIndex = connector.To.ConnectionSite ?? 0;
            if (candidate.Equals(target)) continue;
            source.Wire.Connector = candidate;
            mutations.SemanticChanges = true;
            changedNodes.Add(connector.Id);
            changed = true;
        }
        return changed;
    }

    private static bool IsConnectorFrameLeaf(string kind) => kind is
        "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical" or
        "childLeftEmu" or "childTopEmu" or "childWidthEmu" or "childHeightEmu";

    private static PpjFrameModel ConnectorLeafFrame(PpjFrameModel frame, NativeLeafMutation mutation)
    {
        var value = double.Parse(mutation.After, CultureInfo.InvariantCulture);
        return mutation.LeafKind switch
        {
            "leftEmu" or "childLeftEmu" => frame with { X = value / EmuPerPoint },
            "topEmu" or "childTopEmu" => frame with { Y = value / EmuPerPoint },
            "widthEmu" or "childWidthEmu" => frame with { Width = value / EmuPerPoint },
            "heightEmu" or "childHeightEmu" => frame with { Height = value / EmuPerPoint },
            "rotationDegrees" => frame with { Rotation = value / 60_000 },
            "flipHorizontal" => frame with { FlipH = value != 0 },
            "flipVertical" => frame with { FlipV = value != 0 },
            _ => throw Unsupported(mutation.ProgramElementId, "unknown connector dependency frame leaf"),
        };
    }

    private static void ApplyConnectorFrameLeafToWire(NativeLeafMutation mutation, string path)
    {
        var element = mutation.Element;
        IMessage payload = element.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => element.Shape,
            PresentationElement.ContentOneofCase.Image => element.Image,
            PresentationElement.ContentOneofCase.Group => element.Group,
            PresentationElement.ContentOneofCase.Chart => element.Chart,
            PresentationElement.ContentOneofCase.Table => element.Table,
            _ => throw Unsupported(path, "unsupported connector dependency frame owner"),
        };
        var name = mutation.LeafKind switch
        {
            "leftEmu" => "left_emu", "topEmu" => "top_emu", "widthEmu" => "width_emu", "heightEmu" => "height_emu",
            "childLeftEmu" => "child_left_emu", "childTopEmu" => "child_top_emu",
            "childWidthEmu" => "child_width_emu", "childHeightEmu" => "child_height_emu",
            "rotationDegrees" => "rotation_angle_60000", "flipHorizontal" => "flip_horizontal", "flipVertical" => "flip_vertical",
            _ => throw Unsupported(path, "unknown connector dependency frame field"),
        };
        var transform = mutation.LeafKind is "rotationDegrees" or "flipHorizontal" or "flipVertical";
        if (transform)
        {
            if (element.Shape?.DirectFrame is { } direct) payload = direct;
            else
            {
                var field = payload.Descriptor.FindFieldByName(element.Shape is not null || element.Image is not null ? "transform" : "frame_transform");
                var existing = field.Accessor.GetValue(payload) as IMessage;
                if (existing is null) throw Unsupported(path, "issued transform leaf has no native transform owner");
                payload = existing;
            }
        }
        void Set(IMessage owner)
        {
            var field = owner.Descriptor.FindFieldByName(name) ?? throw Unsupported(path, "issued frame leaf has no modeled field");
            object Value(string text) => field.FieldType switch
            {
                FieldType.Int64 => long.Parse(text, CultureInfo.InvariantCulture),
                FieldType.SInt32 => int.Parse(text, CultureInfo.InvariantCulture),
                FieldType.Bool => text == "1",
                _ => throw Unsupported(path, "unsupported frame scalar"),
            };
            if (!Equals(field.Accessor.GetValue(owner), Value(mutation.Before)))
                throw Unsupported(path, "conflicting semantic and native frame edits");
            field.Accessor.SetValue(owner, Value(mutation.After));
        }
        Set(payload);
        if (!transform && element.Shape?.DirectFrame is { } placeholder) Set(placeholder);
    }
}
