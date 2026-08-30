using System.Text.Json;

namespace OfficeKit.Codec;

internal static class PpjSemanticValidator
{
    private static readonly IReadOnlySet<string> AuthoredVideoMimeTypes = Set("video/mp4");
    private static readonly IReadOnlySet<string> AuthoredAudioMimeTypes = Set("audio/mpeg", "audio/mp4", "audio/wav", "audio/x-wav");

    private static readonly IReadOnlyDictionary<string, IReadOnlySet<string>> CapabilityFields =
        new Dictionary<string, IReadOnlySet<string>>(StringComparer.Ordinal)
        {
            ["replaceText"] = Set("text", "visibleText"),
            ["setFill"] = Set("fill"),
            ["setStroke"] = Set("stroke"),
            ["setOpacity"] = Set("opacity"),
            ["setFrame"] = Set("frame.x", "frame.y", "frame.width", "frame.height", "frame.rotation", "frame.flipH", "frame.flipV"),
            ["setGeometry"] = Set("geometry.adjustments"),
            ["setBackground"] = Set("background"),
            ["replaceImage"] = Set("image.asset"),
            ["setImageCrop"] = Set("image.crop"),
            ["setImageFit"] = Set("image.fit"),
            ["setImageMask"] = Set("image.mask.adjustments"),
            ["setChartTitle"] = Set("chart.title"),
            ["setChartData"] = Set("chart.data"),
            ["setChartTextStyle"] = Set("chart.textStyle"),
            ["setChartFill"] = Set("chart.fill"),
            ["setSmartArtText"] = Set("smartArt.text"),
            ["setOlePayload"] = Set("ole.payload"),
            ["delete"] = Set("element"),
            ["duplicate"] = Set("element"),
            ["reorder"] = Set("zOrder"),
        };

    internal static void Validate(PpjProgramModel program, List<PpjDiagnostic> diagnostics)
    {
        var assets = UniqueIndex(program.Assets, item => item.Id, "$.assets", diagnostics);
        var components = UniqueIndex(program.Components, item => item.Id, "$.components", diagnostics);
        var pages = UniqueIndex(program.Pages, item => item.Id, "$.pages", diagnostics);
        var masters = UniqueIndex(program.Design.Masters, item => item.Id, "$.design.masters", diagnostics);
        var layouts = UniqueIndex(program.Design.Layouts, item => item.Id, "$.design.layouts", diagnostics);

        ValidateUniqueIds(program.Sections, item => item.Id, "$.sections", diagnostics);
        ValidateUniqueIds(program.CustomShows, item => item.Id, "$.customShows", diagnostics);
        ValidateUniqueIds(program.Comments, item => item.Id, "$.comments", diagnostics);

        ValidateRelativeResource(program.Source?.Uri, "$.source.uri", diagnostics);
        for (var index = 0; index < program.Assets.Count; index++)
            ValidateRelativeResource(program.Assets[index].Uri, $"$.assets[{index}].uri", diagnostics);

        var assetIds = assets.Keys.ToHashSet(StringComparer.Ordinal);
        var pageIds = pages.Keys.ToHashSet(StringComparer.Ordinal);
        ValidateResourceReferences(program.Root, "$", assetIds, program.Design.ColorIds, program.Design.FontIds, diagnostics);
        ValidateTextEffects(program.Root, "$", diagnostics);
        ValidateMasterLayoutState(program, masters, layouts, diagnostics);
        ValidateComponentDefinitions(program, components, assetIds, diagnostics);

        var globalElementIds = new HashSet<string>(StringComparer.Ordinal);
        var globalAnimationIds = new HashSet<string>(StringComparer.Ordinal);
        var pageElements = new Dictionary<string, IReadOnlyDictionary<string, PpjElementModel>>(StringComparer.Ordinal);

        for (var pageIndex = 0; pageIndex < program.Pages.Count; pageIndex++)
        {
            var page = program.Pages[pageIndex];
            var pagePath = $"$.pages[{pageIndex}]";
            ValidateNativeRef(page.NativeRef, program.Source, $"{pagePath}.nativeRef", diagnostics);

            var localElements = IndexElements(
                page.Elements,
                $"{pagePath}.elements",
                globalElementIds,
                diagnostics);
            pageElements[page.Id] = localElements;

            foreach (var (element, path) in WalkElements(page.Elements, $"{pagePath}.elements"))
                ValidateElement(element, path, program, components, assetIds, localElements, inComponent: false, diagnostics);

            ValidateAnimations(page, pagePath, localElements, globalAnimationIds, diagnostics);
        }

        ValidateTransitions(program.Pages, pages, pageElements, diagnostics);
        ValidatePresentationReferences(program, pageIds, pageElements, diagnostics);
        ValidateComponentGraph(program.Components, diagnostics);
        ValidateExpansionBudget(program, components, diagnostics);
    }

    private static void ValidateMasterLayoutState(
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjMasterModel> masters,
        IReadOnlyDictionary<string, PpjLayoutModel> layouts,
        List<PpjDiagnostic> diagnostics)
    {
        if (program.Source is null && program.Design.Masters.Count > 1)
            diagnostics.Add(new(
                "ppj.master.count",
                "Source-free PPJ supports exactly one canonical master when master state is declared.",
                "$.design.masters"));
        if (program.Source is null && program.Design.Layouts.Count > 0 && program.Design.Masters.Count != 1)
            diagnostics.Add(new(
                "ppj.layout.master",
                "Source-free PPJ layouts require one declared master.",
                "$.design.layouts"));

        for (var masterIndex = 0; masterIndex < program.Design.Masters.Count; masterIndex++)
        {
            var master = program.Design.Masters[masterIndex];
            var path = $"$.design.masters[{masterIndex}]";
            ValidateTextStyleLevels(master.TitleTextLevels, path + ".textStyles.title", diagnostics);
            ValidateTextStyleLevels(master.BodyTextLevels, path + ".textStyles.body", diagnostics);
            ValidateTextStyleLevels(master.OtherTextLevels, path + ".textStyles.other", diagnostics);
            ValidateLayoutPlaceholders(master.Placeholders, path + ".placeholders", diagnostics);
        }

        for (var layoutIndex = 0; layoutIndex < program.Design.Layouts.Count; layoutIndex++)
        {
            var layout = program.Design.Layouts[layoutIndex];
            var path = $"$.design.layouts[{layoutIndex}]";
            if (!masters.ContainsKey(layout.MasterId))
                diagnostics.Add(new(
                    "ppj.layout.master",
                    $"Layout {layout.Id} references missing master {layout.MasterId}.",
                    path + ".master"));
            ValidateLayoutPlaceholders(layout.Placeholders, path + ".placeholders", diagnostics);
        }

        for (var pageIndex = 0; pageIndex < program.Pages.Count; pageIndex++)
        {
            var page = program.Pages[pageIndex];
            var path = $"$.pages[{pageIndex}]";
            if (page.LayoutId is null)
            {
                if (program.Source is null && program.Design.Layouts.Count > 0)
                    diagnostics.Add(new(
                        "ppj.page.layout",
                        $"Page {page.Id} must bind one declared layout.",
                        path + ".layout"));
                continue;
            }
            if (!layouts.TryGetValue(page.LayoutId, out var layout))
            {
                // A third-party projection records the immutable native layout
                // identity without pretending that its complete graph is
                // authored PPJ state.
                if (program.Source is null)
                    diagnostics.Add(new(
                        "ppj.page.layout",
                        $"Page {page.Id} references missing layout {page.LayoutId}.",
                        path + ".layout"));
                continue;
            }
            if (!masters.TryGetValue(layout.MasterId, out var master)) continue;
            var available = master.Placeholders.Concat(layout.Placeholders)
                .Select(item => (Type: NativePlaceholderType(item.PlaceholderType), item.Index))
                .ToHashSet();
            for (var elementIndex = 0; elementIndex < page.Elements.Count; elementIndex++)
            {
                if (page.Elements[elementIndex] is not PpjPlaceholderElementModel placeholder) continue;
                var elementPath = $"{path}.elements[{elementIndex}]";
                if (!placeholder.Raw.TryGetProperty("index", out var index))
                {
                    diagnostics.Add(new(
                        "ppj.placeholder.index",
                        $"Layout-bound placeholder {placeholder.Id} requires an explicit index.",
                        elementPath + ".index"));
                    continue;
                }
                var identity = (NativePlaceholderType(placeholder.PlaceholderType), index.GetUInt32());
                if (!available.Contains(identity))
                    diagnostics.Add(new(
                        "ppj.placeholder.binding",
                        $"Placeholder {placeholder.Id} has no matching master/layout type and index.",
                        elementPath));
            }
        }
    }

    private static void ValidateTextStyleLevels(
        IReadOnlyList<JsonElement> levels,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var seen = new HashSet<int>();
        for (var index = 0; index < levels.Count; index++)
        {
            var level = levels[index].GetProperty("level").GetInt32();
            if (!seen.Add(level))
                diagnostics.Add(new(
                    "ppj.master.textStyleLevel",
                    $"Master text style level {level} is duplicated.",
                    $"{path}[{index}].level"));
        }
    }

    private static void ValidateLayoutPlaceholders(
        IReadOnlyList<PpjLayoutPlaceholderModel> placeholders,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var identities = new HashSet<(string Type, uint Index)>();
        for (var index = 0; index < placeholders.Count; index++)
        {
            var placeholder = placeholders[index];
            if (!ids.Add(placeholder.Id))
                diagnostics.Add(new(
                    "ppj.placeholder.id",
                    $"Placeholder ID {placeholder.Id} is duplicated within its owner.",
                    $"{path}[{index}].id"));
            var identity = (NativePlaceholderType(placeholder.PlaceholderType), placeholder.Index);
            if (!identities.Add(identity))
                diagnostics.Add(new(
                    "ppj.placeholder.identity",
                    $"Placeholder type/index {placeholder.PlaceholderType}/{placeholder.Index} is duplicated within its owner.",
                    $"{path}[{index}]"));
        }
    }

    private static string NativePlaceholderType(string value) => value switch
    {
        "centered-title" or "centerTitle" => "ctrTitle",
        "subtitle" => "subTitle",
        "content" => "body",
        _ => value,
    };

    private static void ValidateTextEffects(JsonElement value, string path, List<PpjDiagnostic> diagnostics)
    {
        if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
                ValidateTextEffects(item, $"{path}[{index++}]", diagnostics);
            return;
        }
        if (value.ValueKind != JsonValueKind.Object) return;

        var hasColor = value.TryGetProperty("color", out _);
        if (value.TryGetProperty("gradient", out var gradient))
        {
            if (hasColor)
                diagnostics.Add(new("ppj.text.paintConflict", "Text style cannot declare both color and gradient.", path));
            if (gradient.TryGetProperty("kind", out var kind) && kind.GetString() == "radial" && gradient.TryGetProperty("angle", out _))
                diagnostics.Add(new("ppj.text.radialAngle", "Radial text gradients cannot declare a linear angle.", path + ".gradient.angle"));
            if (gradient.TryGetProperty("stops", out var stops))
            {
                var previous = double.NegativeInfinity;
                var stopIndex = 0;
                foreach (var stop in stops.EnumerateArray())
                {
                    var offset = stop.GetProperty("offset").GetDouble();
                    if (offset < previous)
                        diagnostics.Add(new("ppj.text.gradientOrder", "Text gradient stop offsets must be ordered.", $"{path}.gradient.stops[{stopIndex}].offset"));
                    previous = offset;
                    stopIndex++;
                }
            }
        }

        foreach (var property in value.EnumerateObject())
            ValidateTextEffects(property.Value, PpjJsonPath.Property(path, property.Name), diagnostics);
    }

    private static void ValidateComponentDefinitions(
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjComponentModel> components,
        IReadOnlySet<string> assetIds,
        List<PpjDiagnostic> diagnostics)
    {
        for (var componentIndex = 0; componentIndex < program.Components.Count; componentIndex++)
        {
            var component = program.Components[componentIndex];
            var path = $"$.components[{componentIndex}]";
            var parameters = UniqueIndex(component.Parameters, item => item.Name, $"{path}.parameters", diagnostics);
            var slots = UniqueIndex(component.Slots, item => item.Name, $"{path}.slots", diagnostics);
            ValidateUniqueIds(component.Variants, item => item.Id, $"{path}.variants", diagnostics);

            for (var parameterIndex = 0; parameterIndex < component.Parameters.Count; parameterIndex++)
            {
                var parameter = component.Parameters[parameterIndex];
                var parameterPath = $"{path}.parameters[{parameterIndex}]";
                if (parameter.Default is { } defaultValue)
                    ValidateParameterValue(parameter, defaultValue, $"{parameterPath}.default", diagnostics);
                for (var allowedIndex = 0; allowedIndex < parameter.Allowed.Count; allowedIndex++)
                    ValidateParameterValue(parameter, parameter.Allowed[allowedIndex], $"{parameterPath}.allowed[{allowedIndex}]", diagnostics);
                if (parameter.Default is { } candidate && parameter.Allowed.Count > 0 &&
                    !parameter.Allowed.Any(value => JsonEqual(value, candidate)))
                {
                    diagnostics.Add(new(
                        "ppj.component.defaultNotAllowed",
                        $"Default value for parameter {parameter.Name} is not included in its allowed values.",
                        $"{parameterPath}.default"));
                }
            }

            var localIds = new HashSet<string>(StringComparer.Ordinal);
            var localElements = IndexElements(component.Elements, $"{path}.elements", localIds, diagnostics);
            foreach (var (element, elementPath) in WalkElements(component.Elements, $"{path}.elements"))
            {
                if (element.NativeRef is not null)
                    diagnostics.Add(new("ppj.component.nativeRef", "Component definitions cannot clone source-bound nativeRef identities.", $"{elementPath}.nativeRef"));
                ValidateElement(element, elementPath, program, components, assetIds, localElements, inComponent: true, diagnostics);
            }

            for (var bindingIndex = 0; bindingIndex < component.Bindings.Count; bindingIndex++)
            {
                var binding = component.Bindings[bindingIndex];
                var bindingPath = $"{path}.bindings[{bindingIndex}]";
                if (!localElements.TryGetValue(binding.TargetId, out var target))
                    diagnostics.Add(new("ppj.component.unknownBindingTarget", $"Binding target {binding.TargetId} does not exist in component {component.Id}.", $"{bindingPath}.target"));
                else if (!SupportsBinding(target, binding.Field))
                    diagnostics.Add(new("ppj.component.bindingType", $"Field {binding.Field} cannot be assigned on a {target.Type} element.", $"{bindingPath}.field"));
                if (!parameters.ContainsKey(binding.Parameter))
                    diagnostics.Add(new("ppj.component.unknownParameter", $"Binding parameter {binding.Parameter} is not declared by component {component.Id}.", $"{bindingPath}.parameter"));
            }

            for (var variantIndex = 0; variantIndex < component.Variants.Count; variantIndex++)
            {
                var variant = component.Variants[variantIndex];
                var variantPath = $"{path}.variants[{variantIndex}]";
                ValidateCondition(variant.When, parameters, $"{variantPath}.when", diagnostics);
                for (var assignmentIndex = 0; assignmentIndex < variant.Assignments.Count; assignmentIndex++)
                {
                    var assignment = variant.Assignments[assignmentIndex];
                    var assignmentPath = $"{variantPath}.set[{assignmentIndex}]";
                    if (!localElements.TryGetValue(assignment.TargetId, out var target))
                        diagnostics.Add(new("ppj.component.unknownAssignmentTarget", $"Variant target {assignment.TargetId} does not exist in component {component.Id}.", $"{assignmentPath}.target"));
                    else if (!SupportsBinding(target, assignment.Field))
                        diagnostics.Add(new("ppj.component.assignmentType", $"Field {assignment.Field} cannot be assigned on a {target.Type} element.", $"{assignmentPath}.field"));
                }
            }

            foreach (var (element, elementPath) in WalkElements(component.Elements, $"{path}.elements"))
            {
                if (element is PpjSlotElementModel slot && !slots.ContainsKey(slot.SlotId))
                    diagnostics.Add(new("ppj.component.unknownSlot", $"Slot {slot.SlotId} is not declared by component {component.Id}.", $"{elementPath}.slot"));
            }
        }
    }

    private static void ValidateElement(
        PpjElementModel element,
        string path,
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjComponentModel> components,
        IReadOnlySet<string> assetIds,
        IReadOnlyDictionary<string, PpjElementModel> localElements,
        bool inComponent,
        List<PpjDiagnostic> diagnostics)
    {
        ValidateNativeRef(element.NativeRef, program.Source, $"{path}.nativeRef", diagnostics);

        switch (element)
        {
            case PpjTextElementModel text:
                ValidateStyleRef(text.StyleRef, program.Design.TextStyleIds, $"{path}.styleRef", diagnostics);
                break;
            case PpjShapeElementModel shape:
                ValidateStyleRef(shape.StyleRef, program.Design.ShapeStyleIds, $"{path}.styleRef", diagnostics);
                ValidatePresetAdjustments(shape.GeometryKind, shape.GeometryPreset, shape.GeometryAdjustments, path + ".geometry", diagnostics);
                if (shape.GeometryKind == "custom")
                    ValidateCustomGeometry(shape.Raw.GetProperty("geometry"), path + ".geometry", diagnostics);
                break;
            case PpjImageElementModel image:
                ValidateAssetRef(image.AssetId, assetIds, $"{path}.asset", diagnostics);
                if (image.MaskKind is not null)
                    ValidatePresetAdjustments(image.MaskKind, image.MaskPreset, image.MaskAdjustments, path + ".mask", diagnostics);
                if (image.MaskKind == "custom")
                    ValidateCustomGeometry(image.Raw.GetProperty("mask"), path + ".mask", diagnostics);
                break;
            case PpjChartElementModel chart:
                ValidateStyleRef(chart.StyleRef, program.Design.ChartStyleIds, $"{path}.styleRef", diagnostics);
                ValidateChart(chart, path, diagnostics);
                break;
            case PpjTableElementModel table:
                ValidateStyleRef(table.StyleRef, program.Design.TableStyleIds, $"{path}.styleRef", diagnostics);
                ValidateTable(table, path, diagnostics);
                break;
            case PpjConnectorElementModel connector:
                ValidateConnectorEndpoint(connector.From, localElements, $"{path}.from", diagnostics);
                ValidateConnectorEndpoint(connector.To, localElements, $"{path}.to", diagnostics);
                break;
            case PpjMediaElementModel media:
                ValidateAssetRef(media.AssetId, assetIds, $"{path}.asset", diagnostics);
                ValidateAssetRef(media.PosterAssetId, assetIds, $"{path}.posterAsset", diagnostics);
                var mediaAsset = program.Assets.FirstOrDefault(asset => asset.Id.Equals(media.AssetId, StringComparison.Ordinal));
                if (mediaAsset is not null)
                {
                    var supported = media.MediaType == "video"
                        ? AuthoredVideoMimeTypes.Contains(mediaAsset.MimeType)
                        : AuthoredAudioMimeTypes.Contains(mediaAsset.MimeType);
                    if (!supported)
                        diagnostics.Add(new(
                            "ppj.media.mime",
                            $"PPJ {media.MediaType} element {element.Id} cannot compile asset MIME {mediaAsset.MimeType}.",
                            $"{path}.asset"));
                }
                var poster = program.Assets.FirstOrDefault(asset => asset.Id.Equals(media.PosterAssetId, StringComparison.Ordinal));
                if (poster is not null && !poster.MimeType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                    diagnostics.Add(new(
                        "ppj.media.posterMime",
                        $"PPJ media element {element.Id} requires an image poster asset.",
                        $"{path}.posterAsset"));
                break;
            case PpjPlaceholderElementModel placeholder:
                ValidateStyleRef(placeholder.StyleRef, program.Design.TextStyleIds, $"{path}.styleRef", diagnostics);
                break;
            case PpjSmartArtElementModel smartArt:
                ValidateSmartArt(
                    smartArt,
                    path,
                    program.Source,
                    program.Design.TextStyleIds,
                    program.Design.ShapeStyleIds,
                    assetIds,
                    diagnostics);
                break;
            case PpjOleElementModel ole:
                if (ole.PayloadAssetId is not null) ValidateAssetRef(ole.PayloadAssetId, assetIds, $"{path}.payloadAsset", diagnostics);
                if (ole.PreviewAssetId is not null) ValidateAssetRef(ole.PreviewAssetId, assetIds, $"{path}.previewAsset", diagnostics);
                break;
            case PpjOpaqueElementModel opaque:
                if (opaque.NativeRef is null)
                    diagnostics.Add(new("ppj.nativeRef.required", "Opaque elements require a source-bound nativeRef.", $"{path}.nativeRef"));
                if (opaque.PreviewAssetId is not null) ValidateAssetRef(opaque.PreviewAssetId, assetIds, $"{path}.previewAsset", diagnostics);
                break;
            case PpjComponentElementModel instance:
                ValidateComponentInstance(instance, path, components, diagnostics);
                break;
            case PpjSlotElementModel when !inComponent:
                diagnostics.Add(new("ppj.component.slotScope", "Slot placeholders are only valid inside component definitions.", path));
                break;
        }
    }

    private static void ValidatePresetAdjustments(
        string geometryKind,
        string? geometryPreset,
        IReadOnlyList<int> adjustments,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (geometryKind != "preset" || adjustments.Count == 0) return;
        if (geometryPreset is null ||
            !PptxPresetGeometryAdjustmentCodec.TryExpectedCount(geometryPreset, out var expectedCount))
        {
            diagnostics.Add(new(
                "ppj.geometry.adjustmentProfile",
                $"Preset geometry {geometryPreset ?? "(missing)"} has no canonical adjustment profile.",
                path + ".adjustments"));
            return;
        }
        if (adjustments.Count != expectedCount)
            diagnostics.Add(new(
                "ppj.geometry.adjustmentCount",
                $"Preset geometry {geometryPreset} requires either no explicit adjustments or exactly {expectedCount} ordered values.",
                path + ".adjustments"));
    }

    private static void ValidateCustomGeometry(
        JsonElement geometry,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var pathIndex = 0;
        foreach (var geometryPath in geometry.GetProperty("paths").EnumerateArray())
        {
            var hasCurrentPoint = false;
            var hasSubpathStart = false;
            var commandIndex = 0;
            foreach (var command in geometryPath.GetProperty("commands").EnumerateArray())
            {
                var operation = command.GetProperty("op").GetString();
                if (operation == "arcTo" && !hasCurrentPoint)
                    diagnostics.Add(new(
                        "ppj.geometry.arcCurrentPoint",
                        "A custom-geometry arc requires a preceding command that establishes the current point.",
                        $"{path}.paths[{pathIndex}].commands[{commandIndex}]"));

                switch (operation)
                {
                    case "moveTo":
                        hasCurrentPoint = true;
                        hasSubpathStart = true;
                        break;
                    case "lineTo":
                    case "quadraticTo":
                    case "cubicTo":
                    case "arcTo":
                        hasCurrentPoint = true;
                        break;
                    case "close":
                        hasCurrentPoint = hasSubpathStart;
                        break;
                }
                commandIndex++;
            }
            pathIndex++;
        }
    }

    private static void ValidateChart(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        var seriesIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < chart.Data.Series.Count; index++)
        {
            var series = chart.Data.Series[index];
            var seriesPath = $"{path}.data.series[{index}]";
            var seriesType = chart.ChartType == "combo" ? series.ChartType : chart.ChartType;
            if (!seriesIds.Add(series.Id))
                diagnostics.Add(new("ppj.id.duplicate", $"Chart series ID {series.Id} is duplicated.", $"{seriesPath}.id"));
            if (seriesType is not ("scatter" or "bubble" or "sankey") && series.Values.Count != chart.Data.Categories.Count)
                diagnostics.Add(new("ppj.chart.lengthMismatch", $"Series {series.Id} has {series.Values.Count} values for {chart.Data.Categories.Count} categories.", $"{seriesPath}.values"));
            if (chart.ChartType != "waterfall" && series.PointRoles.Count != 0)
                diagnostics.Add(new("ppj.chart.pointRoleType", "pointRoles applies only to waterfall charts.", $"{seriesPath}.pointRoles"));
            if (chart.ChartType == "combo" && string.IsNullOrEmpty(series.ChartType))
                diagnostics.Add(new("ppj.chart.comboSeriesType", "Every combo-chart series requires chartType.", $"{seriesPath}.chartType"));
            if (chart.ChartType != "combo" && series.ChartType is not null && !string.Equals(series.ChartType, chart.ChartType, StringComparison.Ordinal))
                diagnostics.Add(new("ppj.chart.seriesType", "A non-combo chart series cannot override the deck chart type.", $"{seriesPath}.chartType"));
            if (seriesType is "scatter" or "bubble")
            {
                if (chart.Data.Categories.Count != 0)
                    diagnostics.Add(new("ppj.chart.numericCategories", "Scatter and bubble charts require an empty shared categories array.", path + ".data.categories"));
                if (series.XValues.Count != series.Values.Count)
                    diagnostics.Add(new("ppj.chart.xValueLength", $"Series {series.Id} requires one xValue per value.", seriesPath + ".xValues"));
                if (seriesType == "bubble" && series.BubbleSizes.Count != series.Values.Count)
                    diagnostics.Add(new("ppj.chart.bubbleSizeLength", $"Bubble series {series.Id} requires one positive bubbleSize per value.", seriesPath + ".bubbleSizes"));
                if (seriesType == "scatter" && series.BubbleSizes.Count != 0)
                    diagnostics.Add(new("ppj.chart.bubbleSizeType", "bubbleSizes applies only to bubble charts.", seriesPath + ".bubbleSizes"));
            }
            else
            {
                if (series.XValues.Count != 0)
                    diagnostics.Add(new("ppj.chart.xValueType", "xValues applies only to scatter and bubble charts.", seriesPath + ".xValues"));
                if (series.BubbleSizes.Count != 0)
                    diagnostics.Add(new("ppj.chart.bubbleSizeType", "bubbleSizes applies only to bubble charts.", seriesPath + ".bubbleSizes"));
            }
            if (chart.ChartType != "candlestick" &&
                (series.OpenValues.Count != 0 || series.HighValues.Count != 0 || series.LowValues.Count != 0))
                diagnostics.Add(new(
                    "ppj.chart.ohlcType",
                    "openValues, highValues, and lowValues apply only to candlestick charts.",
                    seriesPath));
            if (chart.ChartType is not ("treemap" or "sunburst") && series.Parents.Count != 0)
                diagnostics.Add(new(
                    "ppj.chart.parentsType",
                    "parents applies only to treemap and sunburst charts.",
                    seriesPath + ".parents"));
            if (chart.ChartType != "sankey" && (series.Sources.Count != 0 || series.Targets.Count != 0))
                diagnostics.Add(new(
                    "ppj.chart.edgeChannelType",
                    "sources and targets apply only to sankey charts.",
                    seriesPath));
            if (series.Raw.TryGetProperty("trendlines", out var trendlines))
            {
                var trendlineIndex = 0;
                foreach (var trendline in trendlines.EnumerateArray())
                {
                    var trendlinePath = $"{seriesPath}.trendlines[{trendlineIndex}]";
                    var type = trendline.GetProperty("type").GetString();
                    if (type == "polynomial" && !trendline.TryGetProperty("order", out _))
                        diagnostics.Add(new("ppj.chart.trendlineOrder", "Polynomial trendlines require order.", trendlinePath + ".order"));
                    if (type != "polynomial" && trendline.TryGetProperty("order", out _))
                        diagnostics.Add(new("ppj.chart.trendlineOrder", "order applies only to polynomial trendlines.", trendlinePath + ".order"));
                    if (type == "moving-average" && !trendline.TryGetProperty("period", out _))
                        diagnostics.Add(new("ppj.chart.trendlinePeriod", "Moving-average trendlines require period.", trendlinePath + ".period"));
                    if (type != "moving-average" && trendline.TryGetProperty("period", out _))
                        diagnostics.Add(new("ppj.chart.trendlinePeriod", "period applies only to moving-average trendlines.", trendlinePath + ".period"));
                    trendlineIndex++;
                }
            }
            if (series.Raw.TryGetProperty("errorBars", out var errorBars))
            {
                var valueType = errorBars.GetProperty("valueType").GetString();
                var requiresValue = valueType is "fixed-value" or "percentage" or "standard-deviation";
                if (requiresValue != errorBars.TryGetProperty("value", out _))
                    diagnostics.Add(new(
                        "ppj.chart.errorBarValue",
                        requiresValue ? $"{valueType} error bars require value." : "standard-error error bars do not accept value.",
                        seriesPath + ".errorBars.value"));
            }
        }

        if (chart.ChartType == "waterfall") ValidateWaterfall(chart, path, diagnostics);
        else if (chart.Raw.TryGetProperty("style", out var ordinaryStyle) && ordinaryStyle.TryGetProperty("waterfall", out _))
            diagnostics.Add(new("ppj.chart.waterfallStyleType", "style.waterfall applies only to waterfall charts.", path + ".style.waterfall"));
        if (chart.ChartType == "heatmap") ValidateHeatmap(chart, path, diagnostics);
        else if (chart.Raw.TryGetProperty("style", out var heatmapStyle) && heatmapStyle.TryGetProperty("heatmap", out _))
            diagnostics.Add(new("ppj.chart.heatmapStyleType", "style.heatmap applies only to heatmap charts.", path + ".style.heatmap"));
        if (chart.ChartType == "candlestick") ValidateCandlestick(chart, path, diagnostics);
        else if (chart.Raw.TryGetProperty("style", out var candlestickStyle) && candlestickStyle.TryGetProperty("candlestick", out _))
            diagnostics.Add(new("ppj.chart.candlestickStyleType", "style.candlestick applies only to candlestick charts.", path + ".style.candlestick"));
        if (chart.ChartType == "treemap") ValidateTreemap(chart, path, diagnostics);
        else if (chart.Raw.TryGetProperty("style", out var treemapStyle) && treemapStyle.TryGetProperty("treemap", out _))
            diagnostics.Add(new("ppj.chart.treemapStyleType", "style.treemap applies only to treemap charts.", path + ".style.treemap"));
        if (chart.ChartType == "sunburst") ValidateSunburst(chart, path, diagnostics);
        else if (chart.Raw.TryGetProperty("style", out var sunburstStyle) && sunburstStyle.TryGetProperty("sunburst", out _))
            diagnostics.Add(new("ppj.chart.sunburstStyleType", "style.sunburst applies only to sunburst charts.", path + ".style.sunburst"));
        if (chart.ChartType == "sankey") ValidateSankey(chart, path, diagnostics);
        else if (chart.Raw.TryGetProperty("style", out var sankeyStyle) && sankeyStyle.TryGetProperty("sankey", out _))
            diagnostics.Add(new("ppj.chart.sankeyStyleType", "style.sankey applies only to sankey charts.", path + ".style.sankey"));
        if (chart.ChartType == "combo") ValidateCombo(chart, path, diagnostics);

        if (chart.Raw.TryGetProperty("style", out var style) &&
            style.TryGetProperty("dataLabels", out _) &&
            (style.TryGetProperty("showDataLabels", out _) || style.TryGetProperty("dataLabelPosition", out _)))
            diagnostics.Add(new(
                "ppj.chart.dataLabelConflict",
                "Structured dataLabels cannot be combined with showDataLabels or dataLabelPosition.",
                path + ".style.dataLabels"));
        if (chart.ChartType != "combo" &&
            (chart.Raw.TryGetProperty("secondaryXAxis", out _) || chart.Raw.TryGetProperty("secondaryYAxis", out _)))
            diagnostics.Add(new(
                "ppj.chart.secondaryAxis",
                "Secondary axes are valid only for combo charts.",
                path));
        ValidateAxisBounds(chart.Raw, "xAxis", path, diagnostics);
        ValidateAxisBounds(chart.Raw, "yAxis", path, diagnostics);
        ValidateAxisBounds(chart.Raw, "secondaryXAxis", path, diagnostics);
        ValidateAxisBounds(chart.Raw, "secondaryYAxis", path, diagnostics);
        ValidateAxisKinds(chart.Raw, "xAxis", path, chart.ChartType is not ("scatter" or "bubble"), diagnostics);
        ValidateAxisKinds(chart.Raw, "yAxis", path, categoryAxis: false, diagnostics);
        ValidateAxisKinds(chart.Raw, "secondaryXAxis", path, categoryAxis: true, diagnostics);
        ValidateAxisKinds(chart.Raw, "secondaryYAxis", path, categoryAxis: false, diagnostics);
    }

    private static void ValidateCombo(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        var typed = chart.Data.Series
            .Where(series => !string.IsNullOrEmpty(series.ChartType))
            .ToArray();
        var families = typed
            .Select(series => series.ChartType!)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (families.Length < 2)
            diagnostics.Add(new(
                "ppj.chart.comboFamilies",
                "Combo charts require at least two distinct column, line, or area plot families.",
                path + ".data.series"));

        foreach (var family in families)
        {
            var axes = typed
                .Where(series => string.Equals(series.ChartType, family, StringComparison.Ordinal))
                .Select(series => series.Axis ?? "primary")
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            if (axes.Length > 1)
                diagnostics.Add(new(
                    "ppj.chart.comboFamilyAxis",
                    $"All {family} series in one combo chart must use the same axis pair.",
                    path + ".data.series"));
        }

        var hasSecondary = typed.Any(series => string.Equals(series.Axis, "secondary", StringComparison.Ordinal));
        var hasPrimary = typed.Any(series => !string.Equals(series.Axis, "secondary", StringComparison.Ordinal));
        if (!hasPrimary)
            diagnostics.Add(new(
                "ppj.chart.comboPrimaryAxis",
                "A combo chart requires at least one primary-axis plot family.",
                path + ".data.series"));
        var hasSecondaryXAxis = chart.Raw.TryGetProperty("secondaryXAxis", out _);
        var hasSecondaryYAxis = chart.Raw.TryGetProperty("secondaryYAxis", out _);
        if (hasSecondaryXAxis != hasSecondaryYAxis)
            diagnostics.Add(new(
                "ppj.chart.comboSecondaryAxisPair",
                "Combo secondary axes must be declared as a complete X/Y pair or omitted together.",
                path));
        if (!hasSecondary && (hasSecondaryXAxis || hasSecondaryYAxis))
            diagnostics.Add(new(
                "ppj.chart.comboUnusedSecondaryAxis",
                "Secondary axes require at least one complete secondary plot family.",
                path));
    }

    private static void ValidateHeatmap(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        if (chart.Data.Categories.Count is < 1 or > 32)
            diagnostics.Add(new(
                "ppj.chart.heatmapCategoryCount",
                "Heatmaps require between 1 and 32 x-axis categories.",
                path + ".data.categories"));
        if (chart.Data.Series.Count is < 1 or > 32)
            diagnostics.Add(new(
                "ppj.chart.heatmapSeriesCount",
                "Heatmaps require between 1 and 32 named y-axis series.",
                path + ".data.series"));

        var categories = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < chart.Data.Categories.Count; index++)
        {
            var category = chart.Data.Categories[index];
            if (category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString()))
                diagnostics.Add(new(
                    "ppj.chart.heatmapCategory",
                    "Heatmap categories must be non-empty strings.",
                    $"{path}.data.categories[{index}]"));
            else if (!categories.Add(category.GetString()!))
                diagnostics.Add(new(
                    "ppj.chart.heatmapCategoryDuplicate",
                    $"Heatmap category {category.GetString()} is duplicated.",
                    $"{path}.data.categories[{index}]"));
        }

        var names = new HashSet<string>(StringComparer.Ordinal);
        var hasValue = false;
        for (var index = 0; index < chart.Data.Series.Count; index++)
        {
            var series = chart.Data.Series[index];
            var seriesPath = $"{path}.data.series[{index}]";
            if (string.IsNullOrWhiteSpace(series.Name))
                diagnostics.Add(new("ppj.chart.heatmapSeriesName", "Heatmap series names must be non-empty.", seriesPath + ".name"));
            else if (!names.Add(series.Name))
                diagnostics.Add(new(
                    "ppj.chart.heatmapSeriesNameDuplicate",
                    $"Heatmap series name {series.Name} is duplicated.",
                    seriesPath + ".name"));
            hasValue |= series.Values.Any(value => value is not null);
            foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
                if (series.Raw.TryGetProperty(property, out _))
                    diagnostics.Add(new(
                        "ppj.chart.heatmapSeriesField",
                        $"{property} is not part of the bounded heatmap series profile.",
                        $"{seriesPath}.{property}"));
        }
        if (!hasValue)
            diagnostics.Add(new("ppj.chart.heatmapEmpty", "Heatmaps require at least one numeric value.", path + ".data.series"));

        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.heatmapAxis",
                    "Heatmap axes are generated from matrix labels and do not accept ChartPart axis configuration.",
                    $"{path}.{property}"));

        if (!chart.Raw.TryGetProperty("style", out var style) || !style.TryGetProperty("heatmap", out var heatmap)) return;
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall" })
            if (style.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.heatmapStyleField",
                    $"{property} is not part of the bounded vector heatmap style profile.",
                    $"{path}.style.{property}"));

        var scale = heatmap.TryGetProperty("scale", out var scaleValue) ? scaleValue.GetString() : "linear";
        var colors = heatmap.GetProperty("colors").GetArrayLength();
        if ((scale == "linear" && colors != 2) || (scale == "diverging" && colors != 3))
            diagnostics.Add(new(
                "ppj.chart.heatmapColorCount",
                scale == "diverging" ? "Diverging heatmaps require exactly three colors." : "Linear heatmaps require exactly two colors.",
                path + ".style.heatmap.colors"));

        double? minimum = null;
        double? maximum = null;
        if (heatmap.TryGetProperty("domain", out var domain))
        {
            minimum = domain[0].GetDouble();
            maximum = domain[1].GetDouble();
            if (minimum >= maximum)
                diagnostics.Add(new(
                    "ppj.chart.heatmapDomain",
                    "Heatmap domain minimum must be smaller than its maximum.",
                    path + ".style.heatmap.domain"));
            var effectiveMidpoint = heatmap.TryGetProperty("midpoint", out var configuredMidpoint)
                ? configuredMidpoint.GetDouble()
                : 0;
            if (scale == "diverging" && minimum < maximum && (effectiveMidpoint <= minimum || effectiveMidpoint >= maximum))
                diagnostics.Add(new(
                    "ppj.chart.heatmapMidpoint",
                    "Diverging heatmap midpoint must lie strictly inside the explicit domain.",
                    path + ".style.heatmap.midpoint"));
        }
        if (heatmap.TryGetProperty("midpoint", out var midpoint))
        {
            if (scale != "diverging")
                diagnostics.Add(new(
                    "ppj.chart.heatmapMidpointType",
                    "Heatmap midpoint applies only to a diverging scale.",
                    path + ".style.heatmap.midpoint"));
        }
    }

    private static void ValidateCandlestick(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        if (chart.Data.Categories.Count is < 1 or > 64)
            diagnostics.Add(new(
                "ppj.chart.candlestickCategoryCount",
                "Candlestick charts require between 1 and 64 ordered categories.",
                path + ".data.categories"));
        var categories = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < chart.Data.Categories.Count; index++)
        {
            var category = chart.Data.Categories[index];
            if (category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString()))
                diagnostics.Add(new(
                    "ppj.chart.candlestickCategory",
                    "Candlestick categories must be non-empty strings.",
                    $"{path}.data.categories[{index}]"));
            else if (!categories.Add(category.GetString()!))
                diagnostics.Add(new(
                    "ppj.chart.candlestickCategoryDuplicate",
                    $"Candlestick category {category.GetString()} is duplicated.",
                    $"{path}.data.categories[{index}]"));
        }
        if (chart.Data.Series.Count != 1)
        {
            diagnostics.Add(new(
                "ppj.chart.candlestickSeriesCount",
                "Candlestick charts require exactly one OHLC or HLC series.",
                path + ".data.series"));
            return;
        }

        var series = chart.Data.Series[0];
        var seriesPath = path + ".data.series[0]";
        var count = chart.Data.Categories.Count;
        if (series.Values.Any(value => value is null))
            diagnostics.Add(new(
                "ppj.chart.candlestickMissingClose",
                "Candlestick close values cannot be missing.",
                seriesPath + ".values"));
        if (series.HighValues.Count != count)
            diagnostics.Add(new(
                "ppj.chart.candlestickHighLength",
                "Candlestick highValues must contain one value per category.",
                seriesPath + ".highValues"));
        if (series.LowValues.Count != count)
            diagnostics.Add(new(
                "ppj.chart.candlestickLowLength",
                "Candlestick lowValues must contain one value per category.",
                seriesPath + ".lowValues"));
        if (series.OpenValues.Count != 0 && series.OpenValues.Count != count)
            diagnostics.Add(new(
                "ppj.chart.candlestickOpenLength",
                "Candlestick openValues must be omitted for HLC or contain one value per category for OHLC.",
                seriesPath + ".openValues"));

        if (series.Values.Count == count && series.Values.All(value => value is not null) &&
            series.HighValues.Count == count && series.LowValues.Count == count &&
            (series.OpenValues.Count == 0 || series.OpenValues.Count == count))
        {
            for (var index = 0; index < count; index++)
            {
                var high = series.HighValues[index];
                var low = series.LowValues[index];
                var close = series.Values[index]!.Value;
                if (low > high || close < low || close > high ||
                    (series.OpenValues.Count != 0 && (series.OpenValues[index] < low || series.OpenValues[index] > high)))
                    diagnostics.Add(new(
                        "ppj.chart.candlestickRange",
                        "Each low must not exceed high, and open/close must lie inside that range.",
                        $"{seriesPath}.values[{index}]"));
            }

            var lowest = series.LowValues.Min();
            var highest = series.HighValues.Max();
            if (chart.Raw.TryGetProperty("yAxis", out var yAxis))
            {
                if (yAxis.TryGetProperty("min", out var minimum) && minimum.GetDouble() > lowest)
                    diagnostics.Add(new(
                        "ppj.chart.candlestickAxisClip",
                        "Explicit yAxis.min must not clip the lowest observation.",
                        path + ".yAxis.min"));
                if (yAxis.TryGetProperty("max", out var maximum) && maximum.GetDouble() < highest)
                    diagnostics.Add(new(
                        "ppj.chart.candlestickAxisClip",
                        "Explicit yAxis.max must not clip the highest observation.",
                        path + ".yAxis.max"));
            }
        }

        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.candlestickSeriesField",
                    $"{property} is not part of the bounded candlestick series profile.",
                    $"{seriesPath}.{property}"));

        if (!chart.Raw.TryGetProperty("style", out var style) || !style.TryGetProperty("candlestick", out var candlestick)) return;
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap" })
            if (style.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.candlestickStyleField",
                    $"{property} is not part of the bounded vector candlestick style profile.",
                    $"{path}.style.{property}"));
        foreach (var role in new[] { "up", "down" })
        {
            var fill = candlestick.GetProperty(role).GetProperty("fill");
            if (fill.GetProperty("type").GetString() is "none" or "image")
                diagnostics.Add(new(
                    "ppj.chart.candlestickBodyFill",
                    "Candlestick body fills must be solid or bounded gradients.",
                    $"{path}.style.candlestick.{role}.fill"));
        }
    }

    private static void ValidateTreemap(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        var count = chart.Data.Categories.Count;
        if (count is < 1 or > 128)
            diagnostics.Add(new(
                "ppj.chart.treemapNodeCount",
                "Treemap charts require between 1 and 128 nodes.",
                path + ".data.categories"));
        if (chart.Data.Series.Count != 1)
        {
            diagnostics.Add(new(
                "ppj.chart.treemapSeriesCount",
                "Treemap charts require exactly one hierarchy series.",
                path + ".data.series"));
            return;
        }

        var names = new string?[count];
        var indexes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var index = 0; index < count; index++)
        {
            var category = chart.Data.Categories[index];
            if (category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString()))
                diagnostics.Add(new(
                    "ppj.chart.treemapCategory",
                    "Treemap categories must be non-empty strings.",
                    $"{path}.data.categories[{index}]"));
            else
            {
                names[index] = category.GetString()!;
                if (!indexes.TryAdd(names[index]!, index))
                    diagnostics.Add(new(
                        "ppj.chart.treemapCategoryDuplicate",
                        $"Treemap category {names[index]} is duplicated.",
                        $"{path}.data.categories[{index}]"));
            }
        }

        var series = chart.Data.Series[0];
        var seriesPath = path + ".data.series[0]";
        if (series.Parents.Count != count)
            diagnostics.Add(new(
                "ppj.chart.treemapParentLength",
                "Treemap parents must contain one string or null per category.",
                seriesPath + ".parents"));
        if (series.Values.Any(value => value is null || value <= 0))
            diagnostics.Add(new(
                "ppj.chart.treemapValue",
                "Treemap values must be present and strictly positive.",
                seriesPath + ".values"));

        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.treemapSeriesField",
                    $"{property} is not part of the bounded treemap series profile.",
                    $"{seriesPath}.{property}"));
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.treemapAxis",
                    "Treemap charts do not use Cartesian axes.",
                    $"{path}.{property}"));

        if (series.Parents.Count == count && indexes.Count == count)
        {
            var roots = 0;
            for (var index = 0; index < count; index++)
            {
                var parent = series.Parents[index];
                if (parent is null)
                {
                    roots++;
                    continue;
                }
                if (!indexes.ContainsKey(parent))
                    diagnostics.Add(new(
                        "ppj.chart.treemapMissingParent",
                        $"Treemap parent {parent} does not name a declared category.",
                        $"{seriesPath}.parents[{index}]"));
                else if (string.Equals(parent, names[index], StringComparison.Ordinal))
                    diagnostics.Add(new(
                        "ppj.chart.treemapCycle",
                        "A treemap node cannot parent itself.",
                        $"{seriesPath}.parents[{index}]"));
            }
            if (roots is < 1 or > 16)
                diagnostics.Add(new(
                    "ppj.chart.treemapRootCount",
                    "Treemap charts require between 1 and 16 roots.",
                    seriesPath + ".parents"));

            for (var index = 0; index < count; index++)
            {
                var seen = new HashSet<int>();
                var current = index;
                var depth = 0;
                while (true)
                {
                    if (!seen.Add(current))
                    {
                        diagnostics.Add(new(
                            "ppj.chart.treemapCycle",
                            $"Treemap parent chain for {names[index]} contains a cycle.",
                            $"{seriesPath}.parents[{index}]"));
                        break;
                    }
                    var parent = series.Parents[current];
                    if (parent is null || !indexes.TryGetValue(parent, out current)) break;
                    depth++;
                    if (depth > 8)
                    {
                        diagnostics.Add(new(
                            "ppj.chart.treemapDepth",
                            $"Treemap node {names[index]} exceeds the maximum hierarchy depth of eight.",
                            $"{seriesPath}.parents[{index}]"));
                        break;
                    }
                }
            }

            if (series.Values.Count == count && series.Values.All(value => value is > 0))
            {
                var childSums = new Dictionary<int, double>();
                for (var index = 0; index < count; index++)
                    if (series.Parents[index] is { } parent && indexes.TryGetValue(parent, out var parentIndex))
                        childSums[parentIndex] = childSums.GetValueOrDefault(parentIndex) + series.Values[index]!.Value;
                foreach (var pair in childSums)
                {
                    var declared = series.Values[pair.Key]!.Value;
                    var tolerance = Math.Max(1e-9, Math.Abs(declared) * 1e-9);
                    if (Math.Abs(declared - pair.Value) > tolerance)
                        diagnostics.Add(new(
                            "ppj.chart.treemapTotal",
                            $"Treemap parent {names[pair.Key]} value {declared} does not equal its direct-child sum {pair.Value}.",
                            $"{seriesPath}.values[{pair.Key}]"));
                }
            }
        }

        if (!chart.Raw.TryGetProperty("style", out var style) || !style.TryGetProperty("treemap", out _)) return;
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick" })
            if (style.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.treemapStyleField",
                    $"{property} is not part of the bounded vector treemap style profile.",
                    $"{path}.style.{property}"));
    }

    private static void ValidateSunburst(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        var count = chart.Data.Categories.Count;
        if (count is < 1 or > 96)
            diagnostics.Add(new(
                "ppj.chart.sunburstNodeCount",
                "Sunburst charts require between 1 and 96 nodes.",
                path + ".data.categories"));
        if (chart.Data.Series.Count != 1)
        {
            diagnostics.Add(new(
                "ppj.chart.sunburstSeriesCount",
                "Sunburst charts require exactly one hierarchy series.",
                path + ".data.series"));
            return;
        }

        var names = new string?[count];
        var indexes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var index = 0; index < count; index++)
        {
            var category = chart.Data.Categories[index];
            if (category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString()))
                diagnostics.Add(new(
                    "ppj.chart.sunburstCategory",
                    "Sunburst categories must be non-empty strings.",
                    $"{path}.data.categories[{index}]"));
            else
            {
                names[index] = category.GetString()!;
                if (!indexes.TryAdd(names[index]!, index))
                    diagnostics.Add(new(
                        "ppj.chart.sunburstCategoryDuplicate",
                        $"Sunburst category {names[index]} is duplicated.",
                        $"{path}.data.categories[{index}]"));
            }
        }

        var series = chart.Data.Series[0];
        var seriesPath = path + ".data.series[0]";
        if (series.Values.Count != count)
            diagnostics.Add(new(
                "ppj.chart.sunburstValueLength",
                "Sunburst values must contain one number per category.",
                seriesPath + ".values"));
        if (series.Parents.Count != count)
            diagnostics.Add(new(
                "ppj.chart.sunburstParentLength",
                "Sunburst parents must contain one string or null per category.",
                seriesPath + ".parents"));
        if (series.Values.Any(value => value is null || value <= 0))
            diagnostics.Add(new(
                "ppj.chart.sunburstValue",
                "Sunburst values must be present and strictly positive.",
                seriesPath + ".values"));

        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.sunburstSeriesField",
                    $"{property} is not part of the bounded sunburst series profile.",
                    $"{seriesPath}.{property}"));
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.sunburstAxis",
                    "Sunburst charts do not use Cartesian axes.",
                    $"{path}.{property}"));

        if (series.Parents.Count == count && series.Values.Count == count && indexes.Count == count)
        {
            var roots = 0;
            for (var index = 0; index < count; index++)
            {
                var parent = series.Parents[index];
                if (parent is null)
                {
                    roots++;
                    continue;
                }
                if (!indexes.ContainsKey(parent))
                    diagnostics.Add(new(
                        "ppj.chart.sunburstMissingParent",
                        $"Sunburst parent {parent} does not name a declared category.",
                        $"{seriesPath}.parents[{index}]"));
                else if (string.Equals(parent, names[index], StringComparison.Ordinal))
                    diagnostics.Add(new(
                        "ppj.chart.sunburstCycle",
                        "A sunburst node cannot parent itself.",
                        $"{seriesPath}.parents[{index}]"));
            }
            if (roots is < 1 or > 16)
                diagnostics.Add(new(
                    "ppj.chart.sunburstRootCount",
                    "Sunburst charts require between 1 and 16 roots.",
                    seriesPath + ".parents"));

            for (var index = 0; index < count; index++)
            {
                var seen = new HashSet<int>();
                var current = index;
                var depth = 0;
                while (true)
                {
                    if (!seen.Add(current))
                    {
                        diagnostics.Add(new(
                            "ppj.chart.sunburstCycle",
                            $"Sunburst parent chain for {names[index]} contains a cycle.",
                            $"{seriesPath}.parents[{index}]"));
                        break;
                    }
                    var parent = series.Parents[current];
                    if (parent is null || !indexes.TryGetValue(parent, out current)) break;
                    depth++;
                    if (depth > 6)
                    {
                        diagnostics.Add(new(
                            "ppj.chart.sunburstDepth",
                            $"Sunburst node {names[index]} exceeds the maximum hierarchy depth of six.",
                            $"{seriesPath}.parents[{index}]"));
                        break;
                    }
                }
            }

            if (series.Values.All(value => value is > 0))
            {
                var childSums = new Dictionary<int, double>();
                for (var index = 0; index < count; index++)
                    if (series.Parents[index] is { } parent && indexes.TryGetValue(parent, out var parentIndex))
                        childSums[parentIndex] = childSums.GetValueOrDefault(parentIndex) + series.Values[index]!.Value;
                foreach (var pair in childSums)
                {
                    var declared = series.Values[pair.Key]!.Value;
                    var tolerance = Math.Max(1e-9, Math.Abs(declared) * 1e-9);
                    if (Math.Abs(declared - pair.Value) > tolerance)
                        diagnostics.Add(new(
                            "ppj.chart.sunburstTotal",
                            $"Sunburst parent {names[pair.Key]} value {declared} does not equal its direct-child sum {pair.Value}.",
                            $"{seriesPath}.values[{pair.Key}]"));
                }
            }
        }

        if (!chart.Raw.TryGetProperty("style", out var style) || !style.TryGetProperty("sunburst", out _)) return;
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap" })
            if (style.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.sunburstStyleField",
                    $"{property} is not part of the bounded vector sunburst style profile.",
                    $"{path}.style.{property}"));
    }

    private static void ValidateSankey(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        var nodeCount = chart.Data.Categories.Count;
        if (nodeCount is < 2 or > 64)
            diagnostics.Add(new(
                "ppj.chart.sankeyNodeCount",
                "Sankey charts require between 2 and 64 declared nodes.",
                path + ".data.categories"));
        if (chart.Data.Series.Count != 1)
        {
            diagnostics.Add(new(
                "ppj.chart.sankeySeriesCount",
                "Sankey charts require exactly one directed-flow series.",
                path + ".data.series"));
            return;
        }

        var nodes = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < nodeCount; index++)
        {
            var category = chart.Data.Categories[index];
            if (category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString()))
                diagnostics.Add(new(
                    "ppj.chart.sankeyNode",
                    "Sankey node names must be non-empty strings.",
                    $"{path}.data.categories[{index}]"));
            else if (!nodes.Add(category.GetString()!))
                diagnostics.Add(new(
                    "ppj.chart.sankeyNodeDuplicate",
                    $"Sankey node {category.GetString()} is duplicated.",
                    $"{path}.data.categories[{index}]"));
        }

        var series = chart.Data.Series[0];
        var seriesPath = path + ".data.series[0]";
        var edgeCount = series.Values.Count;
        if (edgeCount is < 1 or > 256)
            diagnostics.Add(new(
                "ppj.chart.sankeyEdgeCount",
                "Sankey charts require between 1 and 256 directed edges.",
                seriesPath + ".values"));
        if (series.Sources.Count != edgeCount)
            diagnostics.Add(new(
                "ppj.chart.sankeySourceLength",
                "Sankey sources must contain one node name per flow value.",
                seriesPath + ".sources"));
        if (series.Targets.Count != edgeCount)
            diagnostics.Add(new(
                "ppj.chart.sankeyTargetLength",
                "Sankey targets must contain one node name per flow value.",
                seriesPath + ".targets"));
        if (series.Values.Any(value => value is null || value <= 0))
            diagnostics.Add(new(
                "ppj.chart.sankeyFlow",
                "Sankey flow values must be present and strictly positive.",
                seriesPath + ".values"));

        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "parents", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.sankeySeriesField",
                    $"{property} is not part of the bounded sankey series profile.",
                    $"{seriesPath}.{property}"));
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.sankeyAxis",
                    "Sankey charts do not use Cartesian axes.",
                    $"{path}.{property}"));

        if (nodes.Count == nodeCount && series.Sources.Count == edgeCount && series.Targets.Count == edgeCount)
        {
            var indegree = nodes.ToDictionary(node => node, _ => 0, StringComparer.Ordinal);
            var outgoing = nodes.ToDictionary(node => node, _ => new List<string>(), StringComparer.Ordinal);
            var incomingFlow = nodes.ToDictionary(node => node, _ => 0d, StringComparer.Ordinal);
            var outgoingFlow = nodes.ToDictionary(node => node, _ => 0d, StringComparer.Ordinal);
            var used = new HashSet<string>(StringComparer.Ordinal);
            var edges = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 0; index < edgeCount; index++)
            {
                var source = series.Sources[index];
                var target = series.Targets[index];
                if (!nodes.Contains(source))
                    diagnostics.Add(new(
                        "ppj.chart.sankeyMissingNode",
                        $"Sankey source {source} is not a declared node.",
                        $"{seriesPath}.sources[{index}]"));
                if (!nodes.Contains(target))
                    diagnostics.Add(new(
                        "ppj.chart.sankeyMissingNode",
                        $"Sankey target {target} is not a declared node.",
                        $"{seriesPath}.targets[{index}]"));
                if (!nodes.Contains(source) || !nodes.Contains(target)) continue;
                if (string.Equals(source, target, StringComparison.Ordinal))
                {
                    diagnostics.Add(new(
                        "ppj.chart.sankeyCycle",
                        "A Sankey edge cannot target its source node.",
                        $"{seriesPath}.targets[{index}]"));
                    continue;
                }
                if (!edges.Add(source + "\u0000" + target))
                    diagnostics.Add(new(
                        "ppj.chart.sankeyEdgeDuplicate",
                        $"Sankey edge {source} to {target} is duplicated; combine its flow value explicitly.",
                        $"{seriesPath}.targets[{index}]"));
                outgoing[source].Add(target);
                indegree[target]++;
                used.Add(source);
                used.Add(target);
                if (series.Values[index] is { } flow)
                {
                    outgoingFlow[source] += flow;
                    incomingFlow[target] += flow;
                }
            }

            foreach (var node in nodes.Where(node => !used.Contains(node)))
                diagnostics.Add(new(
                    "ppj.chart.sankeyDisconnectedNode",
                    $"Sankey node {node} is not incident to any declared edge.",
                    path + ".data.categories"));

            var queue = new Queue<string>(chart.Data.Categories
                .Select(category => category.GetString())
                .Where(node => node is not null && indegree[node] == 0)!
                .Select(node => node!));
            var visited = 0;
            while (queue.Count != 0)
            {
                var node = queue.Dequeue();
                visited++;
                foreach (var target in outgoing[node])
                    if (--indegree[target] == 0) queue.Enqueue(target);
            }
            if (visited != nodes.Count)
                diagnostics.Add(new(
                    "ppj.chart.sankeyCycle",
                    "Sankey edges must form a directed acyclic graph.",
                    seriesPath));

            foreach (var node in nodes.Where(node => incomingFlow[node] > 0 && outgoingFlow[node] > 0))
            {
                var tolerance = Math.Max(1e-9, Math.Max(incomingFlow[node], outgoingFlow[node]) * 1e-9);
                if (Math.Abs(incomingFlow[node] - outgoingFlow[node]) > tolerance)
                    diagnostics.Add(new(
                        "ppj.chart.sankeyConservation",
                        $"Sankey internal node {node} has incoming flow {incomingFlow[node]} and outgoing flow {outgoingFlow[node]}.",
                        seriesPath + ".values"));
            }
        }

        if (!chart.Raw.TryGetProperty("style", out var style) || !style.TryGetProperty("sankey", out _)) return;
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap", "sunburst" })
            if (style.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.sankeyStyleField",
                    $"{property} is not part of the bounded vector sankey style profile.",
                    $"{path}.style.{property}"));
    }

    private static void ValidateWaterfall(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        if (chart.Data.Series.Count != 1)
        {
            diagnostics.Add(new(
                "ppj.chart.waterfallSeriesCount",
                "Waterfall charts require exactly one semantic series.",
                path + ".data.series"));
            return;
        }

        var series = chart.Data.Series[0];
        var seriesPath = path + ".data.series[0]";
        if (series.PointRoles.Count != series.Values.Count)
            diagnostics.Add(new(
                "ppj.chart.waterfallRoleLength",
                "Waterfall pointRoles must contain exactly one delta or total role per value.",
                seriesPath + ".pointRoles"));
        if (series.Values.Any(value => value is null))
            diagnostics.Add(new(
                "ppj.chart.waterfallMissingValue",
                "Waterfall values cannot be missing.",
                seriesPath + ".values"));

        foreach (var property in new[] { "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars", "xValues", "bubbleSizes" })
            if (series.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.waterfallSeriesField",
                    $"{property} is not part of the bounded waterfall series profile.",
                    $"{seriesPath}.{property}"));

        if (series.PointRoles.Count == series.Values.Count && series.Values.All(value => value is not null))
        {
            const double tolerance = 1e-9;
            var running = 0d;
            for (var index = 0; index < series.Values.Count; index++)
            {
                var value = series.Values[index]!.Value;
                if (series.PointRoles[index] == "total")
                {
                    if (value < -tolerance)
                        diagnostics.Add(new(
                            "ppj.chart.waterfallNegativeTotal",
                            "Waterfall total values must be non-negative in the bounded profile.",
                            $"{seriesPath}.values[{index}]"));
                    if (index > 0 && Math.Abs(value - running) > tolerance)
                        diagnostics.Add(new(
                            "ppj.chart.waterfallTotalMismatch",
                            $"Waterfall total {value.ToString(System.Globalization.CultureInfo.InvariantCulture)} does not equal the computed running total {running.ToString(System.Globalization.CultureInfo.InvariantCulture)}.",
                            $"{seriesPath}.values[{index}]"));
                    running = value;
                }
                else
                {
                    running += value;
                    if (running < -tolerance)
                        diagnostics.Add(new(
                            "ppj.chart.waterfallNegativeRunningTotal",
                            "Waterfall cumulative values cannot cross below zero in the bounded profile.",
                            $"{seriesPath}.values[{index}]"));
                }
            }
        }

        if (chart.Raw.TryGetProperty("secondaryXAxis", out _) || chart.Raw.TryGetProperty("secondaryYAxis", out _))
            diagnostics.Add(new("ppj.chart.waterfallSecondaryAxis", "Waterfall charts do not support secondary axes.", path));
        if (chart.Raw.TryGetProperty("yAxis", out var yAxis) &&
            yAxis.TryGetProperty("min", out var minimum) && minimum.GetDouble() > 0)
            diagnostics.Add(new(
                "ppj.chart.waterfallAxisMinimum",
                "Waterfall value-axis minimum must include zero.",
                path + ".yAxis.min"));

        if (!chart.Raw.TryGetProperty("style", out var style)) return;
        foreach (var property in new[] { "stacking", "showDataLabels", "dataLabelPosition", "dataLabels", "smooth", "varyColors" })
            if (style.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.waterfallStyleField",
                    $"{property} is not part of the bounded waterfall style profile.",
                    $"{path}.style.{property}"));
        if (style.TryGetProperty("legend", out var legend) && legend.GetString() != "none")
            diagnostics.Add(new(
                "ppj.chart.waterfallLegend",
                "Waterfall charts do not expose the internal lowering series through a legend.",
                path + ".style.legend"));
        if (!style.TryGetProperty("waterfall", out var waterfall)) return;
        foreach (var role in new[] { "increase", "decrease", "total" })
        {
            var roleFill = waterfall.GetProperty(role).GetProperty("fill");
            if (roleFill.GetProperty("type").GetString() is "none" or "image")
                diagnostics.Add(new(
                    "ppj.chart.waterfallRoleFill",
                    "Waterfall role fills must be solid or bounded gradients.",
                    $"{path}.style.waterfall.{role}.fill"));
        }
    }

    private static void ValidateAxisKinds(
        JsonElement chart,
        string property,
        string path,
        bool categoryAxis,
        List<PpjDiagnostic> diagnostics)
    {
        if (!chart.TryGetProperty(property, out var axis)) return;
        if (categoryAxis)
        {
            foreach (var name in new[] { "min", "max", "majorUnit" })
                if (axis.TryGetProperty(name, out _))
                    diagnostics.Add(new(
                        "ppj.chart.axisField",
                        $"{name} applies only to a value axis.",
                        $"{path}.{property}.{name}"));
        }
        else if (axis.TryGetProperty("tickLabelInterval", out _))
            diagnostics.Add(new(
                "ppj.chart.axisField",
                "tickLabelInterval applies only to a category axis.",
                $"{path}.{property}.tickLabelInterval"));
    }

    private static void ValidateAxisBounds(
        JsonElement chart,
        string property,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (!chart.TryGetProperty(property, out var axis) ||
            !axis.TryGetProperty("min", out var minimum) ||
            !axis.TryGetProperty("max", out var maximum)) return;
        if (minimum.GetDouble() >= maximum.GetDouble())
            diagnostics.Add(new(
                "ppj.chart.axisBounds",
                "Axis min must be less than max.",
                $"{path}.{property}"));
    }

    private static void ValidateTable(PpjTableElementModel table, string path, List<PpjDiagnostic> diagnostics)
    {
        var columnIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < table.Columns.Count; index++)
        {
            var id = table.Columns[index].Id;
            if (id is not null && !columnIds.Add(id))
                diagnostics.Add(new("ppj.id.duplicate", $"Table column ID {id} is duplicated.", $"{path}.columns[{index}].id"));
        }

        var rowIds = new HashSet<string>(StringComparer.Ordinal);
        var cellIds = new HashSet<string>(StringComparer.Ordinal);
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            if (row.Id is not null && !rowIds.Add(row.Id))
                diagnostics.Add(new("ppj.id.duplicate", $"Table row ID {row.Id} is duplicated.", $"{path}.rows[{rowIndex}].id"));
            var usedColumns = 0;
            for (var cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                var cell = row.Cells[cellIndex];
                if (cell.Id is not null && !cellIds.Add(cell.Id))
                    diagnostics.Add(new("ppj.id.duplicate", $"Table cell ID {cell.Id} is duplicated.", $"{path}.rows[{rowIndex}].cells[{cellIndex}].id"));
                usedColumns += cell.ColumnSpan;
                if (rowIndex + cell.RowSpan > table.Rows.Count)
                    diagnostics.Add(new("ppj.table.rowSpan", "Cell rowSpan extends beyond the table.", $"{path}.rows[{rowIndex}].cells[{cellIndex}].rowSpan"));
            }
            if (usedColumns != table.Columns.Count)
                diagnostics.Add(new("ppj.table.columnSpan", $"Row spans {usedColumns} columns but the table declares {table.Columns.Count}.", $"{path}.rows[{rowIndex}].cells"));
        }
    }

    private static void ValidateSmartArt(
        PpjSmartArtElementModel smartArt,
        string path,
        PpjSourceModel? source,
        IReadOnlySet<string> textStyleIds,
        IReadOnlySet<string> shapeStyleIds,
        IReadOnlySet<string> assetIds,
        List<PpjDiagnostic> diagnostics)
    {
        if (smartArt.Mode == "authored" && smartArt.Layout is null)
            diagnostics.Add(new("ppj.smartArt.layout", "Authored SmartArt requires a supported layout.", $"{path}.layout"));
        if (smartArt.Mode == "source-bound" && smartArt.NativeRef is null && smartArt.Nodes.All(node => node.NativeRef is null))
            diagnostics.Add(new("ppj.smartArt.nativeRef", "Source-bound SmartArt requires an element or node nativeRef.", $"{path}.nativeRef"));

        if (smartArt.Mode == "authored")
        {
            if (smartArt.NativeRef is not null || smartArt.Nodes.Any(node => node.NativeRef is not null))
                diagnostics.Add(new("ppj.smartArt.nativeRef", "Authored diagrams cannot carry source-bound nativeRef authority.", path));
            if (smartArt.Nodes.Count is < 1 or > 64)
                diagnostics.Add(new("ppj.smartArt.nodeCount", "Authored diagrams require between 1 and 64 nodes.", $"{path}.nodes"));
            if (smartArt.ShapeStyleRef is null)
                diagnostics.Add(new("ppj.smartArt.shapeStyle", "Authored diagrams require an explicit default shapeStyleRef.", $"{path}.shapeStyleRef"));
            if (smartArt.TextStyleRef is null)
                diagnostics.Add(new("ppj.smartArt.textStyle", "Authored diagrams require an explicit default textStyleRef.", $"{path}.textStyleRef"));
            if (smartArt.ShapeStyleRef is not null)
                ValidateStyleRef(smartArt.ShapeStyleRef, shapeStyleIds, $"{path}.shapeStyleRef", diagnostics);
            if (smartArt.TextStyleRef is not null)
                ValidateStyleRef(smartArt.TextStyleRef, textStyleIds, $"{path}.textStyleRef", diagnostics);
            if (smartArt.Raw.TryGetProperty("nodeGeometry", out var nodeGeometry))
                ValidateDiagramGeometry(nodeGeometry, $"{path}.nodeGeometry", diagnostics);

            var hasParentEdges = smartArt.Nodes.Any(node => node.ParentId is not null);
            if (hasParentEdges && smartArt.Layout is ("list" or "process" or "cycle" or "matrix" or "pyramid" or "picture"))
                diagnostics.Add(new("ppj.smartArt.topology", $"Authored {smartArt.Layout} diagrams use ordered nodes and cannot declare parent edges.", $"{path}.nodes"));
            if (smartArt.Layout == "hierarchy" && smartArt.Nodes.Count > 1 && !hasParentEdges)
                diagnostics.Add(new("ppj.smartArt.topology", "Authored hierarchy diagrams require parent edges.", $"{path}.nodes"));
            if (smartArt.Nodes.Count > 1 && smartArt.Layout is ("process" or "cycle" or "hierarchy" or "relationship") &&
                !smartArt.Raw.TryGetProperty("connector", out _))
                diagnostics.Add(new("ppj.smartArt.connector", $"Authored {smartArt.Layout} diagrams require explicit connector styling.", $"{path}.connector"));
            if ((smartArt.Nodes.Count < 2 || smartArt.Layout is ("list" or "matrix" or "pyramid" or "picture")) &&
                smartArt.Raw.TryGetProperty("connector", out _))
                diagnostics.Add(new("ppj.smartArt.connector", $"Authored {smartArt.Layout} diagrams do not emit connector edges for this node set.", $"{path}.connector"));
        }

        var nodes = UniqueIndex(smartArt.Nodes, node => node.Id, $"{path}.nodes", diagnostics);
        var rawNodes = smartArt.Raw.GetProperty("nodes").EnumerateArray().ToArray();
        for (var index = 0; index < smartArt.Nodes.Count; index++)
        {
            var node = smartArt.Nodes[index];
            var nodePath = $"{path}.nodes[{index}]";
            if (node.ParentId is not null && !nodes.ContainsKey(node.ParentId))
                diagnostics.Add(new("ppj.smartArt.parent", $"SmartArt parent {node.ParentId} does not exist.", $"{nodePath}.parent"));
            ValidateStyleRef(node.StyleRef, textStyleIds, $"{nodePath}.styleRef", diagnostics);
            ValidateStyleRef(node.ShapeStyleRef, shapeStyleIds, $"{nodePath}.shapeStyleRef", diagnostics);
            if (rawNodes[index].TryGetProperty("geometry", out var geometry))
                ValidateDiagramGeometry(geometry, $"{nodePath}.geometry", diagnostics);
            if (node.AssetId is not null)
                ValidateAssetRef(node.AssetId, assetIds, $"{nodePath}.asset", diagnostics);
            if (smartArt.Mode == "authored" && smartArt.Layout == "picture" && node.AssetId is null)
                diagnostics.Add(new("ppj.smartArt.pictureAsset", "Every authored picture-diagram node requires an image asset.", $"{nodePath}.asset"));
            if (smartArt.Mode == "authored" && smartArt.Layout != "picture" && node.AssetId is not null)
                diagnostics.Add(new("ppj.smartArt.pictureAsset", "Diagram node assets are only valid for the picture layout.", $"{nodePath}.asset"));
            ValidateNativeRef(node.NativeRef, source, $"{nodePath}.nativeRef", diagnostics);
        }

        foreach (var node in smartArt.Nodes)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var cursor = node;
            while (cursor.ParentId is not null && nodes.TryGetValue(cursor.ParentId, out cursor!))
            {
                if (!seen.Add(cursor.Id))
                {
                    diagnostics.Add(new("ppj.smartArt.cycle", "SmartArt parent references form a cycle.", $"{path}.nodes"));
                    break;
                }
            }
        }
    }

    private static void ValidateDiagramGeometry(
        JsonElement geometry,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var kind = geometry.GetProperty("kind").GetString()!;
        var preset = geometry.TryGetProperty("preset", out var presetValue) ? presetValue.GetString() : null;
        var adjustments = geometry.TryGetProperty("adjustments", out var values)
            ? values.EnumerateArray().Select(value => value.GetInt32()).ToArray()
            : [];
        ValidatePresetAdjustments(kind, preset, adjustments, path, diagnostics);
        if (kind == "custom") ValidateCustomGeometry(geometry, path, diagnostics);
    }

    private static void ValidateComponentInstance(
        PpjComponentElementModel instance,
        string path,
        IReadOnlyDictionary<string, PpjComponentModel> components,
        List<PpjDiagnostic> diagnostics)
    {
        if (!components.TryGetValue(instance.ComponentId, out var component))
        {
            diagnostics.Add(new("ppj.component.unknown", $"Component {instance.ComponentId} is not declared.", $"{path}.component"));
            return;
        }

        var parameters = component.Parameters
            .GroupBy(item => item.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        ValidateArguments(instance.Arguments, parameters, $"{path}.arguments", diagnostics);
        if (instance.VariantId is not null && component.Variants.All(item => item.Id != instance.VariantId))
            diagnostics.Add(new("ppj.component.unknownVariant", $"Variant {instance.VariantId} is not declared by component {component.Id}.", $"{path}.variant"));
        if (instance.When is not null)
            ValidateCondition(instance.When, parameters, $"{path}.when", diagnostics);

        var slots = component.Slots
            .GroupBy(item => item.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        foreach (var supplied in instance.Slots)
        {
            if (!slots.TryGetValue(supplied.Key, out var slot))
            {
                diagnostics.Add(new("ppj.component.unknownSlot", $"Slot {supplied.Key} is not declared by component {component.Id}.", $"{path}.slots.{supplied.Key}"));
                continue;
            }
            if (supplied.Value.Count < slot.Minimum || supplied.Value.Count > slot.Maximum)
                diagnostics.Add(new("ppj.component.slotCount", $"Slot {supplied.Key} received {supplied.Value.Count} elements; expected {slot.Minimum}..{slot.Maximum}.", $"{path}.slots.{supplied.Key}"));
            for (var index = 0; index < supplied.Value.Count; index++)
            {
                if (!slot.Accepts.Contains(supplied.Value[index].Type, StringComparer.Ordinal))
                    diagnostics.Add(new("ppj.component.slotType", $"Slot {supplied.Key} does not accept {supplied.Value[index].Type} elements.", $"{path}.slots.{supplied.Key}[{index}].type"));
            }
        }
        foreach (var slot in component.Slots)
        {
            var count = instance.Slots.TryGetValue(slot.Name, out var values) ? values.Count : 0;
            if (count < slot.Minimum)
                diagnostics.Add(new("ppj.component.slotRequired", $"Slot {slot.Name} requires at least {slot.Minimum} element(s).", $"{path}.slots"));
        }

        if (instance.Repeat is not null)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 0; index < instance.Repeat.Items.Count; index++)
            {
                var item = instance.Repeat.Items[index];
                if (!keys.Add(item.Key))
                    diagnostics.Add(new("ppj.component.repeatKey", $"Repeat key {item.Key} is duplicated.", $"{path}.repeat.items[{index}].key"));
                var merged = new Dictionary<string, JsonElement>(instance.Arguments, StringComparer.Ordinal);
                foreach (var argument in item.Arguments) merged[argument.Key] = argument.Value;
                ValidateArguments(merged, parameters, $"{path}.repeat.items[{index}].arguments", diagnostics);
                ValidateRequiredArguments(component, merged, $"{path}.repeat.items[{index}].arguments", diagnostics);
            }
        }
        else
        {
            ValidateRequiredArguments(component, instance.Arguments, $"{path}.arguments", diagnostics);
        }
    }

    private static void ValidateRequiredArguments(
        PpjComponentModel component,
        IReadOnlyDictionary<string, JsonElement> arguments,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        foreach (var parameter in component.Parameters.Where(item => item.Required && item.Default is null))
        {
            if (!arguments.ContainsKey(parameter.Name))
                diagnostics.Add(new("ppj.component.requiredArgument", $"Component {component.Id} requires argument {parameter.Name}.", path));
        }
    }

    private static void ValidateArguments(
        IReadOnlyDictionary<string, JsonElement> arguments,
        IReadOnlyDictionary<string, PpjComponentParameterModel> parameters,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        foreach (var argument in arguments)
        {
            if (!parameters.TryGetValue(argument.Key, out var parameter))
            {
                diagnostics.Add(new("ppj.component.unknownArgument", $"Argument {argument.Key} is not declared.", PpjJsonPath.Property(path, argument.Key)));
                continue;
            }
            ValidateParameterValue(parameter, argument.Value, PpjJsonPath.Property(path, argument.Key), diagnostics);
        }
    }

    private static void ValidateParameterValue(PpjComponentParameterModel parameter, JsonElement value, string path, List<PpjDiagnostic> diagnostics)
    {
        var matches = parameter.Type switch
        {
            "string" => value.ValueKind == JsonValueKind.String,
            "number" => value.ValueKind == JsonValueKind.Number,
            "integer" => value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out _),
            "boolean" => value.ValueKind is JsonValueKind.True or JsonValueKind.False,
            "color" => IsWrappedValue(value, "color"),
            "asset" => IsWrappedValue(value, "asset"),
            "text" => value.ValueKind == JsonValueKind.String || IsWrappedValue(value, "text"),
            _ => false,
        };
        if (!matches)
            diagnostics.Add(new("ppj.component.argumentType", $"Parameter {parameter.Name} expects {parameter.Type}.", path));
        if (parameter.Allowed.Count > 0 && !parameter.Allowed.Any(candidate => JsonEqual(candidate, value)))
            diagnostics.Add(new("ppj.component.argumentNotAllowed", $"Value for parameter {parameter.Name} is not allowed.", path));
    }

    private static void ValidateCondition(
        PpjConditionModel condition,
        IReadOnlyDictionary<string, PpjComponentParameterModel> parameters,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (!parameters.ContainsKey(condition.Argument))
            diagnostics.Add(new("ppj.component.conditionArgument", $"Condition argument {condition.Argument} is not declared.", $"{path}.argument"));
        var requiresValue = condition.Operator is "equals" or "notEquals";
        if (requiresValue && condition.Value is null)
            diagnostics.Add(new("ppj.component.conditionValue", $"Operator {condition.Operator} requires a value.", $"{path}.value"));
        if (!requiresValue && condition.Value is not null)
            diagnostics.Add(new("ppj.component.conditionValue", $"Operator {condition.Operator} does not accept a value.", $"{path}.value"));
    }

    private static void ValidateAnimations(
        PpjPageModel page,
        string pagePath,
        IReadOnlyDictionary<string, PpjElementModel> elements,
        HashSet<string> globalAnimationIds,
        List<PpjDiagnostic> diagnostics)
    {
        var expandedTimingNodes = 0;
        for (var index = 0; index < page.Animations.Count; index++)
        {
            var animation = page.Animations[index];
            var path = $"{pagePath}.animations[{index}]";
            if (!globalAnimationIds.Add(animation.Id))
                diagnostics.Add(new("ppj.id.duplicate", $"Animation ID {animation.Id} is duplicated.", $"{path}.id"));
            if (!elements.TryGetValue(animation.TargetId, out var target))
            {
                diagnostics.Add(new("ppj.animation.target", $"Animation target {animation.TargetId} does not exist on page {page.Id}.", $"{path}.target"));
                continue;
            }
            if (animation.TextBuild is not null && !HasText(target))
                diagnostics.Add(new("ppj.animation.textBuild", "textBuild requires a text-bearing target.", $"{path}.textBuild"));
            if (animation.ChartBuild is not null && target is not PpjChartElementModel)
                diagnostics.Add(new("ppj.animation.chartBuild", "chartBuild requires a chart target.", $"{path}.chartBuild"));
            if (animation.ChartBuild is not null && target is PpjChartElementModel { ChartType: "heatmap" or "candlestick" or "treemap" or "sunburst" or "sankey" })
                diagnostics.Add(new(
                    "ppj.animation.vectorChartBuild",
                    "Vector-lowered charts compile to one editable group and support whole-object animation, not ChartPart build modes.",
                    $"{path}.chartBuild"));
            if ((animation.Effect == "pulse") != (animation.Phase == "emphasis"))
                diagnostics.Add(new("ppj.animation.phaseEffect", "pulse is the only emphasis effect and is only valid in the emphasis phase.", $"{path}.effect"));

            expandedTimingNodes += EstimateTimingNodes(animation, target);
        }
        if (expandedTimingNodes > 64)
            diagnostics.Add(new("ppj.animation.timingBudget", $"Page expands to {expandedTimingNodes} timing nodes; the limit is 64.", $"{pagePath}.animations"));
    }

    private static int EstimateTimingNodes(PpjAnimationModel animation, PpjElementModel target)
    {
        if (animation.TextBuild == "paragraph")
        {
            var text = target switch
            {
                PpjTextElementModel item => item.Text,
                PpjShapeElementModel item => item.Text,
                PpjPlaceholderElementModel item => item.Text,
                _ => null,
            };
            return Math.Max(1, text?.Paragraphs.Count ?? 1);
        }
        if (animation.ChartBuild is "series" or "series-element" && target is PpjChartElementModel seriesChart)
            return Math.Max(1, seriesChart.Data.Series.Count);
        if (animation.ChartBuild is "category" or "category-element" && target is PpjChartElementModel categoryChart)
            return Math.Max(1, categoryChart.Data.Categories.Count);
        return 1;
    }

    private static void ValidateTransitions(
        IReadOnlyList<PpjPageModel> pages,
        IReadOnlyDictionary<string, PpjPageModel> pagesById,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, PpjElementModel>> pageElements,
        List<PpjDiagnostic> diagnostics)
    {
        for (var index = 0; index < pages.Count; index++)
        {
            var page = pages[index];
            var transition = page.Transition;
            if (transition is null) continue;
            var path = $"$.pages[{index}].transition";
            if (transition.Type != "morph")
            {
                if (transition.FromPageId is not null || transition.MorphPairs.Count > 0)
                    diagnostics.Add(new("ppj.transition.morphFields", "fromPage and morphPairs are only valid for morph transitions.", path));
                continue;
            }

            if (index == 0 || transition.FromPageId is null || transition.FromPageId != pages[index - 1].Id)
            {
                diagnostics.Add(new("ppj.transition.morphAdjacency", "Morph must name the immediately preceding page as fromPage.", $"{path}.fromPage"));
                continue;
            }
            if (!pagesById.ContainsKey(transition.FromPageId)) continue;
            var sourceElements = pageElements[transition.FromPageId];
            var destinationElements = pageElements[page.Id];
            var keys = new HashSet<string>(StringComparer.Ordinal);
            var fromIds = new HashSet<string>(StringComparer.Ordinal);
            var toIds = new HashSet<string>(StringComparer.Ordinal);
            for (var pairIndex = 0; pairIndex < transition.MorphPairs.Count; pairIndex++)
            {
                var pair = transition.MorphPairs[pairIndex];
                var pairPath = $"{path}.morphPairs[{pairIndex}]";
                if (!keys.Add(pair.Key)) diagnostics.Add(new("ppj.transition.morphKey", $"Morph key {pair.Key} is duplicated.", $"{pairPath}.key"));
                if (!fromIds.Add(pair.FromElementId)) diagnostics.Add(new("ppj.transition.morphFrom", $"Morph source {pair.FromElementId} is paired more than once.", $"{pairPath}.from"));
                if (!toIds.Add(pair.ToElementId)) diagnostics.Add(new("ppj.transition.morphTo", $"Morph destination {pair.ToElementId} is paired more than once.", $"{pairPath}.to"));
                if (!sourceElements.TryGetValue(pair.FromElementId, out var from))
                    diagnostics.Add(new("ppj.transition.morphFrom", $"Morph source {pair.FromElementId} does not exist on {transition.FromPageId}.", $"{pairPath}.from"));
                if (!destinationElements.TryGetValue(pair.ToElementId, out var to))
                    diagnostics.Add(new("ppj.transition.morphTo", $"Morph destination {pair.ToElementId} does not exist on {page.Id}.", $"{pairPath}.to"));
                if (from is not null && to is not null && !MorphCompatible(from, to))
                    diagnostics.Add(new("ppj.transition.morphType", $"Morph cannot pair {from.Type} with {to.Type}.", pairPath));
            }
        }
    }

    private static void ValidatePresentationReferences(
        PpjProgramModel program,
        IReadOnlySet<string> pageIds,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, PpjElementModel>> pageElements,
        List<PpjDiagnostic> diagnostics)
    {
        for (var index = 0; index < program.Sections.Count; index++)
            ValidatePageList(program.Sections[index].PageIds, pageIds, $"$.sections[{index}].pages", diagnostics);
        for (var index = 0; index < program.CustomShows.Count; index++)
            ValidatePageList(program.CustomShows[index].PageIds, pageIds, $"$.customShows[{index}].pages", diagnostics);

        var comments = program.Comments
            .GroupBy(item => item.Id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        for (var index = 0; index < program.Comments.Count; index++)
        {
            var comment = program.Comments[index];
            var path = $"$.comments[{index}]";
            if (!pageIds.Contains(comment.PageId))
                diagnostics.Add(new("ppj.comment.page", $"Comment page {comment.PageId} does not exist.", $"{path}.page"));
            else if (comment.TargetId is not null && !pageElements[comment.PageId].ContainsKey(comment.TargetId))
                diagnostics.Add(new("ppj.comment.target", $"Comment target {comment.TargetId} does not exist on page {comment.PageId}.", $"{path}.target"));
            if (comment.ParentId is not null && !comments.ContainsKey(comment.ParentId))
                diagnostics.Add(new("ppj.comment.parent", $"Parent comment {comment.ParentId} does not exist.", $"{path}.parent"));
            ValidateNativeRef(comment.NativeRef, program.Source, $"{path}.nativeRef", diagnostics);
        }
    }

    private static void ValidateComponentGraph(IReadOnlyList<PpjComponentModel> components, List<PpjDiagnostic> diagnostics)
    {
        var byId = components
            .GroupBy(item => item.Id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var state = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var component in components)
            Visit(component.Id, [], 1);

        void Visit(string id, List<string> stack, int depth)
        {
            if (!byId.TryGetValue(id, out var component)) return;
            if (state.TryGetValue(id, out var current))
            {
                if (current == 1)
                {
                    var cycle = string.Join(" -> ", stack.Append(id));
                    diagnostics.Add(new("ppj.component.cycle", $"Component dependency cycle: {cycle}.", "$.components"));
                }
                return;
            }
            if (depth > PpjProgramValidator.MaxComponentDepth)
            {
                diagnostics.Add(new("ppj.component.depth", $"Component expansion exceeds depth {PpjProgramValidator.MaxComponentDepth}.", "$.components"));
                return;
            }
            state[id] = 1;
            stack.Add(id);
            foreach (var nested in WalkElements(component.Elements, string.Empty).Select(item => item.Element).OfType<PpjComponentElementModel>())
                Visit(nested.ComponentId, stack, depth + 1);
            stack.RemoveAt(stack.Count - 1);
            state[id] = 2;
        }
    }

    private static void ValidateExpansionBudget(
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjComponentModel> components,
        List<PpjDiagnostic> diagnostics)
    {
        long count = 0;
        foreach (var page in program.Pages)
        {
            foreach (var element in page.Elements)
                count = SaturatingAdd(count, EstimateElement(element, components, 1));
        }
        if (count > PpjProgramValidator.MaxExpandedElements)
            diagnostics.Add(new("ppj.component.elementBudget", $"PPJ can expand to {count} elements; the limit is {PpjProgramValidator.MaxExpandedElements}.", "$.pages"));
    }

    private static long EstimateElement(PpjElementModel element, IReadOnlyDictionary<string, PpjComponentModel> components, int depth)
    {
        if (depth > PpjProgramValidator.MaxComponentDepth) return PpjProgramValidator.MaxExpandedElements + 1L;
        if (element is PpjGroupElementModel group)
            return 1 + group.Elements.Aggregate(0L, (value, child) => SaturatingAdd(value, EstimateElement(child, components, depth)));
        if (element is not PpjComponentElementModel instance || !components.TryGetValue(instance.ComponentId, out var component))
            return element is PpjSlotElementModel ? 0 : 1;

        long body = 0;
        foreach (var child in component.Elements)
            body = SaturatingAdd(body, EstimateElement(child, components, depth + 1));
        foreach (var supplied in instance.Slots.Values)
            foreach (var child in supplied)
                body = SaturatingAdd(body, EstimateElement(child, components, depth + 1));
        var multiplier = instance.Repeat?.Items.Count ?? 1;
        return Math.Min(PpjProgramValidator.MaxExpandedElements + 1L, body * multiplier);
    }

    private static long SaturatingAdd(long left, long right) =>
        Math.Min(PpjProgramValidator.MaxExpandedElements + 1L, left + right);

    private static IReadOnlyDictionary<string, PpjElementModel> IndexElements(
        IReadOnlyList<PpjElementModel> elements,
        string path,
        HashSet<string> globalIds,
        List<PpjDiagnostic> diagnostics)
    {
        var result = new Dictionary<string, PpjElementModel>(StringComparer.Ordinal);
        foreach (var (element, elementPath) in WalkElements(elements, path))
        {
            if (!result.TryAdd(element.Id, element))
                diagnostics.Add(new("ppj.id.duplicate", $"Element ID {element.Id} is duplicated in this scope.", $"{elementPath}.id"));
            if (!globalIds.Add(element.Id))
                diagnostics.Add(new("ppj.id.duplicate", $"Element ID {element.Id} is not globally unique.", $"{elementPath}.id"));
        }
        return result;
    }

    private static IEnumerable<(PpjElementModel Element, string Path)> WalkElements(IReadOnlyList<PpjElementModel> elements, string path)
    {
        for (var index = 0; index < elements.Count; index++)
        {
            var element = elements[index];
            var elementPath = string.IsNullOrEmpty(path) ? string.Empty : $"{path}[{index}]";
            yield return (element, elementPath);
            if (element is PpjGroupElementModel group)
            {
                foreach (var descendant in WalkElements(group.Elements, $"{elementPath}.elements"))
                    yield return descendant;
            }
            if (element is PpjComponentElementModel component)
            {
                foreach (var slot in component.Slots)
                {
                    foreach (var descendant in WalkElements(slot.Value, $"{elementPath}.slots.{slot.Key}"))
                        yield return descendant;
                }
            }
        }
    }

    private static void ValidateNativeRef(PpjNativeRefModel? nativeRef, PpjSourceModel? source, string path, List<PpjDiagnostic> diagnostics)
    {
        if (nativeRef is null) return;
        if (source is null)
        {
            diagnostics.Add(new("ppj.nativeRef.source", "nativeRef is only valid in a source-bound PPJ.", path));
            return;
        }
        if (nativeRef.SourceSha256 != source.Sha256)
            diagnostics.Add(new("ppj.nativeRef.sourceHash", "nativeRef sourceSha256 does not match source.sha256.", $"{path}.sourceSha256"));
        if (nativeRef.Revision != source.Revision)
            diagnostics.Add(new("ppj.nativeRef.revision", "nativeRef revision does not match source.revision.", $"{path}.revision"));

        var capabilityIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < nativeRef.Capabilities.Count; index++)
        {
            var capability = nativeRef.Capabilities[index];
            var capabilityPath = $"{path}.capabilities[{index}]";
            if (!capabilityIds.Add(capability.Id))
                diagnostics.Add(new("ppj.nativeRef.capabilityId", $"Capability ID {capability.Id} is duplicated.", $"{capabilityPath}.id"));
            if (!CapabilityFields.TryGetValue(capability.Operation, out var permitted)) continue;
            foreach (var field in capability.Fields)
            {
                if (!permitted.Contains(field))
                    diagnostics.Add(new("ppj.nativeRef.capabilityField", $"Operation {capability.Operation} cannot issue field {field}.", $"{capabilityPath}.fields"));
            }
        }
        var leafIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < nativeRef.Leaves.Count; index++)
        {
            var leaf = nativeRef.Leaves[index];
            var leafPath = $"{path}.leaves[{index}]";
            if (!leafIds.Add(leaf.Id))
                diagnostics.Add(new("ppj.nativeRef.leafId", $"Native leaf ID {leaf.Id} is duplicated.", $"{leafPath}.id"));
            try
            {
                _ = PpjNativeLeafProjection.NormalizeValue(leaf.Kind, leaf.Value, $"{leafPath}.value");
            }
            catch (CodecException error)
            {
                diagnostics.Add(new(error.Code, error.Message, error.SourcePath ?? $"{leafPath}.value"));
            }
        }
    }

    private static void ValidateConnectorEndpoint(
        PpjConnectorEndpointModel endpoint,
        IReadOnlyDictionary<string, PpjElementModel> elements,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (endpoint.ElementId is not null && !elements.ContainsKey(endpoint.ElementId))
            diagnostics.Add(new("ppj.connector.target", $"Connector endpoint {endpoint.ElementId} does not exist in the same page or component scope.", $"{path}.element"));
    }

    private static void ValidatePageList(IReadOnlyList<string> ids, IReadOnlySet<string> pages, string path, List<PpjDiagnostic> diagnostics)
    {
        for (var index = 0; index < ids.Count; index++)
        {
            if (!pages.Contains(ids[index]))
                diagnostics.Add(new("ppj.pageRef", $"Page {ids[index]} does not exist.", $"{path}[{index}]"));
        }
    }

    private static void ValidateResourceReferences(
        JsonElement value,
        string path,
        IReadOnlySet<string> assetIds,
        IReadOnlySet<string> colorIds,
        IReadOnlySet<string> fontIds,
        List<PpjDiagnostic> diagnostics)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            if (value.TryGetProperty("token", out var token) && token.ValueKind == JsonValueKind.String && !colorIds.Contains(token.GetString()!))
                diagnostics.Add(new("ppj.colorRef", $"Color token {token.GetString()} does not exist.", $"{path}.token"));
            if (value.TryGetProperty("font", out var font) && font.ValueKind == JsonValueKind.String && !fontIds.Contains(font.GetString()!))
                diagnostics.Add(new("ppj.fontRef", $"Font {font.GetString()} does not exist.", $"{path}.font"));
            foreach (var property in value.EnumerateObject())
            {
                var propertyPath = PpjJsonPath.Property(path, property.Name);
                if (property.Name is "asset" or "posterAsset" or "payloadAsset" or "previewAsset" &&
                    property.Value.ValueKind == JsonValueKind.String && !assetIds.Contains(property.Value.GetString()!))
                {
                    diagnostics.Add(new("ppj.assetRef", $"Asset {property.Value.GetString()} does not exist.", propertyPath));
                }
                ValidateResourceReferences(property.Value, propertyPath, assetIds, colorIds, fontIds, diagnostics);
            }
        }
        else if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
                ValidateResourceReferences(item, $"{path}[{index++}]", assetIds, colorIds, fontIds, diagnostics);
        }
    }

    private static void ValidateStyleRef(string? id, IReadOnlySet<string> ids, string path, List<PpjDiagnostic> diagnostics)
    {
        if (id is not null && !ids.Contains(id))
            diagnostics.Add(new("ppj.styleRef", $"Style {id} does not exist in the matching style catalog.", path));
    }

    private static void ValidateAssetRef(string id, IReadOnlySet<string> assets, string path, List<PpjDiagnostic> diagnostics)
    {
        if (!assets.Contains(id)) diagnostics.Add(new("ppj.assetRef", $"Asset {id} does not exist.", path));
    }

    private static void ValidateRelativeResource(string? uri, string path, List<PpjDiagnostic> diagnostics)
    {
        if (uri is null) return;
        var normalized = uri.Replace('\\', '/');
        if (normalized.StartsWith("/", StringComparison.Ordinal) || Uri.TryCreate(normalized, UriKind.Absolute, out _) ||
            normalized.Split('/').Any(segment => segment == ".."))
        {
            diagnostics.Add(new("ppj.relativeUri", "Resource URI must stay relative to the PPJ directory.", path));
        }
    }

    private static bool HasText(PpjElementModel element) => element switch
    {
        PpjTextElementModel => true,
        PpjShapeElementModel shape => shape.Text is not null,
        PpjPlaceholderElementModel placeholder => placeholder.Text is not null,
        _ => false,
    };

    private static bool MorphCompatible(PpjElementModel from, PpjElementModel to) =>
        from is not PpjChartElementModel && to is not PpjChartElementModel && from.Type == to.Type;

    private static bool SupportsBinding(PpjElementModel target, string field) => field switch
    {
        "text" => HasText(target),
        "fill" or "stroke" or "opacity" => target is PpjShapeElementModel or PpjTextElementModel,
        "frame.x" or "frame.y" or "frame.width" or "frame.height" => true,
        "image.asset" or "image.crop" => target is PpjImageElementModel,
        "chart.title" or "chart.data" => target is PpjChartElementModel,
        "table.cell.text" => target is PpjTableElementModel,
        "accessibility.description" => true,
        _ => false,
    };

    private static bool IsWrappedValue(JsonElement value, string property) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(property, out _);

    private static bool JsonEqual(JsonElement left, JsonElement right) =>
        PpjCanonicalJson.Write(left).AsSpan().SequenceEqual(PpjCanonicalJson.Write(right));

    private static IReadOnlyDictionary<string, T> UniqueIndex<T>(
        IReadOnlyList<T> items,
        Func<T, string> id,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var result = new Dictionary<string, T>(StringComparer.Ordinal);
        for (var index = 0; index < items.Count; index++)
        {
            var value = id(items[index]);
            if (!result.TryAdd(value, items[index]))
                diagnostics.Add(new("ppj.id.duplicate", $"ID {value} is duplicated.", $"{path}[{index}].id"));
        }
        return result;
    }

    private static void ValidateUniqueIds<T>(IReadOnlyList<T> items, Func<T, string> id, string path, List<PpjDiagnostic> diagnostics) =>
        UniqueIndex(items, id, path, diagnostics);

    private static IReadOnlySet<string> Set(params string[] values) =>
        values.ToHashSet(StringComparer.Ordinal);
}
