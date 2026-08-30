using System.Text.Json;

namespace OfficeKit.Codec;

internal sealed record PpjProgramModel(
    JsonElement Root,
    PpjMetaModel Meta,
    PpjIntentModel Intent,
    PpjDesignModel Design,
    IReadOnlyList<PpjAssetModel> Assets,
    PpjSourceModel? Source,
    IReadOnlyList<PpjComponentModel> Components,
    IReadOnlyList<PpjPageModel> Pages,
    IReadOnlyList<PpjSectionModel> Sections,
    IReadOnlyList<PpjCustomShowModel> CustomShows,
    IReadOnlyList<PpjCommentModel> Comments);

internal sealed record PpjMetaModel(string Id, string Title, string Language, int Version);

internal sealed record PpjIntentModel(
    string PrimaryJob,
    string ExpectedOutcome,
    string Audience,
    string Thesis,
    string DeliveryMode);

internal sealed record PpjDesignModel(
    double Width,
    double Height,
    PpjNativeRefModel? CanvasNativeRef,
    IReadOnlySet<string> ColorIds,
    IReadOnlySet<string> FontIds,
    IReadOnlySet<string> TextStyleIds,
    IReadOnlySet<string> ShapeStyleIds,
    IReadOnlySet<string> ChartStyleIds,
    IReadOnlySet<string> TableStyleIds,
    string MotionPolicy,
    IReadOnlyList<PpjMasterModel> Masters,
    IReadOnlyList<PpjLayoutModel> Layouts);

internal sealed record PpjMasterModel(
    string Id,
    string Name,
    JsonElement? Background,
    IReadOnlyList<JsonElement> TitleTextLevels,
    IReadOnlyList<JsonElement> BodyTextLevels,
    IReadOnlyList<JsonElement> OtherTextLevels,
    IReadOnlyList<PpjLayoutPlaceholderModel> Placeholders,
    JsonElement Raw);

internal sealed record PpjLayoutModel(
    string Id,
    string Name,
    string MasterId,
    string LayoutType,
    JsonElement? Background,
    IReadOnlyList<PpjLayoutPlaceholderModel> Placeholders,
    JsonElement Raw);

internal sealed record PpjLayoutPlaceholderModel(
    string Id,
    string Name,
    string PlaceholderType,
    uint Index,
    PpjFrameModel Frame,
    PpjTextContentModel? Text,
    JsonElement? Style,
    JsonElement Raw);

internal sealed record PpjAssetModel(
    string Id,
    string Uri,
    string MimeType,
    string Sha256,
    JsonElement Rights,
    JsonElement Accessibility);

internal sealed record PpjSourceModel(
    string Kind,
    string Uri,
    string Sha256,
    string Revision,
    string ProjectionSha256,
    int VisibleObjectCount);

internal sealed record PpjFrameModel(
    double X,
    double Y,
    double Width,
    double Height,
    double Rotation,
    bool FlipH,
    bool FlipV);

internal sealed record PpjAccessibilityModel(bool Decorative, string? Title, string? Description);

internal sealed record PpjNativeCapabilityModel(
    string Id,
    string Operation,
    string ExpectedHash,
    IReadOnlyList<string> Fields);

internal sealed record PpjNativeLeafModel(
    string Id,
    string Kind,
    string ExpectedHash,
    JsonElement Value);

internal sealed record PpjNativeRefModel(
    string Handle,
    string SourceSha256,
    string Revision,
    string ObjectHash,
    string CapabilitySetSha256,
    IReadOnlyList<PpjNativeCapabilityModel> Capabilities,
    IReadOnlyList<PpjNativeLeafModel> Leaves);

internal abstract class PpjElementModel
{
    internal string Id { get; set; } = string.Empty;
    internal string Type { get; set; } = string.Empty;
    internal PpjFrameModel Frame { get; set; } = new(0, 0, 1, 1, 0, false, false);
    internal string? Name { get; set; }
    internal string? Role { get; set; }
    internal PpjAccessibilityModel? Accessibility { get; set; }
    internal bool? Hidden { get; set; }
    internal bool? Locked { get; set; }
    internal PpjNativeRefModel? NativeRef { get; set; }
    internal JsonElement Raw { get; set; }
}

internal sealed class PpjTextElementModel : PpjElementModel
{
    internal required PpjTextContentModel Text { get; init; }
    internal string? StyleRef { get; init; }
}

internal sealed class PpjShapeElementModel : PpjElementModel
{
    internal required string GeometryKind { get; init; }
    internal string? GeometryPreset { get; init; }
    internal IReadOnlyList<int> GeometryAdjustments { get; init; } = [];
    internal PpjTextContentModel? Text { get; init; }
    internal string? StyleRef { get; init; }
}

internal sealed class PpjIconElementModel : PpjElementModel
{
    internal required string IconName { get; init; }
    internal string? StyleRef { get; init; }
}

internal sealed class PpjImageElementModel : PpjElementModel
{
    internal required string AssetId { get; init; }
    internal string? SvgAssetId { get; init; }
    internal string? Fit { get; init; }
    internal string? MaskKind { get; init; }
    internal string? MaskPreset { get; init; }
    internal IReadOnlyList<int> MaskAdjustments { get; init; } = [];
}

internal sealed class PpjChartElementModel : PpjElementModel
{
    internal required string ChartType { get; init; }
    internal PpjTextContentModel? Title { get; init; }
    internal required PpjChartDataModel Data { get; init; }
    internal string? StyleRef { get; init; }
}

internal sealed record PpjChartDataModel(
    IReadOnlyList<JsonElement> Categories,
    IReadOnlyList<PpjChartSeriesModel> Series);

internal sealed record PpjChartSeriesModel(
    string Id,
    string Name,
    IReadOnlyList<double?> Values,
    IReadOnlyList<string> PointRoles,
    IReadOnlyList<double> XValues,
    IReadOnlyList<double> BubbleSizes,
    IReadOnlyList<double> OpenValues,
    IReadOnlyList<double> HighValues,
    IReadOnlyList<double> LowValues,
    IReadOnlyList<string?> Parents,
    IReadOnlyList<string> Sources,
    IReadOnlyList<string> Targets,
    int? Levels,
    string? ChartType,
    string? Axis,
    JsonElement Raw);

internal sealed class PpjTableElementModel : PpjElementModel
{
    internal required IReadOnlyList<PpjTableColumnModel> Columns { get; init; }
    internal required IReadOnlyList<PpjTableRowModel> Rows { get; init; }
    internal string? StyleRef { get; init; }
}

internal sealed record PpjTableColumnModel(string? Id, double Width);

internal sealed record PpjTableRowModel(string? Id, double? Height, IReadOnlyList<PpjTableCellModel> Cells);

internal sealed record PpjTableCellModel(
    string? Id,
    PpjTextContentModel Text,
    int RowSpan,
    int ColumnSpan);

internal sealed class PpjConnectorElementModel : PpjElementModel
{
    internal required string ConnectorType { get; init; }
    internal required PpjConnectorEndpointModel From { get; init; }
    internal required PpjConnectorEndpointModel To { get; init; }
}

internal sealed record PpjConnectorEndpointModel(
    string? ElementId,
    string? Anchor,
    double? X,
    double? Y);

internal sealed class PpjGroupElementModel : PpjElementModel
{
    internal required IReadOnlyList<PpjElementModel> Elements { get; init; }
}

internal sealed class PpjMediaElementModel : PpjElementModel
{
    internal required string MediaType { get; init; }
    internal required string AssetId { get; init; }
    internal required string PosterAssetId { get; init; }
    internal ulong? StartAtMs { get; init; }
    internal ulong? EndAtMs { get; init; }
    internal bool? Loop { get; init; }
    internal bool? Mute { get; init; }
}

internal sealed class PpjPlaceholderElementModel : PpjElementModel
{
    internal required string PlaceholderType { get; init; }
    internal PpjTextContentModel? Text { get; init; }
    internal string? StyleRef { get; init; }
}

internal sealed class PpjSmartArtElementModel : PpjElementModel
{
    internal required string Mode { get; init; }
    internal string? Layout { get; init; }
    internal string? ShapeStyleRef { get; init; }
    internal string? TextStyleRef { get; init; }
    internal required IReadOnlyList<PpjSmartArtNodeModel> Nodes { get; init; }
}

internal sealed record PpjSmartArtNodeModel(
    string Id,
    string? ParentId,
    PpjTextContentModel Text,
    string? StyleRef,
    string? ShapeStyleRef,
    string? AssetId,
    PpjNativeRefModel? NativeRef,
    JsonElement Raw);

internal sealed class PpjOleElementModel : PpjElementModel
{
    internal string? ProgramId { get; init; }
    internal string? PayloadAssetId { get; init; }
    internal string? PreviewAssetId { get; init; }
}

internal sealed class PpjOpaqueElementModel : PpjElementModel
{
    internal required string NativeKind { get; init; }
    internal required string Summary { get; init; }
    internal required IReadOnlyList<string> VisibleText { get; init; }
    internal string? PreviewAssetId { get; init; }
}

internal sealed class PpjComponentElementModel : PpjElementModel
{
    internal required string ComponentId { get; init; }
    internal string? VariantId { get; init; }
    internal required IReadOnlyDictionary<string, JsonElement> Arguments { get; init; }
    internal required IReadOnlyDictionary<string, IReadOnlyList<PpjElementModel>> Slots { get; init; }
    internal PpjRepeatModel? Repeat { get; init; }
    internal PpjConditionModel? When { get; init; }
}

internal sealed class PpjSlotElementModel : PpjElementModel
{
    internal required string SlotId { get; init; }
}

internal sealed record PpjTextContentModel(
    string? PlainText,
    IReadOnlyList<PpjParagraphModel> Paragraphs);

internal sealed record PpjParagraphModel(
    string? Id,
    IReadOnlyList<PpjRunModel> Runs);

internal sealed record PpjFormulaModel(string Syntax, string Source);

internal sealed record PpjRunModel(string? Id, string? Text, PpjFormulaModel? Formula);

internal sealed record PpjPageModel(
    string Id,
    string Role,
    string? Name,
    string? Claim,
    string? LayoutId,
    PpjTextContentModel? Notes,
    bool? Hidden,
    IReadOnlyList<PpjElementModel> Elements,
    IReadOnlyList<PpjAnimationModel> Animations,
    PpjTransitionModel? Transition,
    PpjSourceCloneModel? SourceClone,
    PpjNativeRefModel? NativeRef,
    JsonElement Raw);

internal sealed record PpjSourceCloneModel(
    string PageId,
    string CapabilityId,
    string? RetainElementId);

internal sealed record PpjAnimationModel(
    string Id,
    string TargetId,
    string Phase,
    string Effect,
    string? Direction,
    string Start,
    int DurationMs,
    int DelayMs,
    string? TextBuild,
    string? ChartBuild,
    int StaggerMs,
    bool? AnimateChartBackground);

internal sealed record PpjTransitionModel(
    string Type,
    int? DurationMs,
    string? Direction,
    string? Orientation,
    string? Speed,
    bool? ThroughBlack,
    int? Spokes,
    bool? AdvanceOnClick,
    int? AdvanceAfterMs,
    string? FromPageId,
    IReadOnlyList<PpjMorphPairModel> MorphPairs);

internal sealed record PpjMorphPairModel(string Key, string FromElementId, string ToElementId);

internal sealed record PpjComponentModel(
    string Id,
    PpjFrameModel Frame,
    IReadOnlyList<PpjComponentParameterModel> Parameters,
    IReadOnlyList<PpjSlotDefinitionModel> Slots,
    IReadOnlyList<PpjComponentBindingModel> Bindings,
    IReadOnlyList<PpjComponentVariantModel> Variants,
    IReadOnlyList<PpjElementModel> Elements,
    JsonElement Raw);

internal sealed record PpjComponentParameterModel(
    string Name,
    string Type,
    bool Required,
    JsonElement? Default,
    IReadOnlyList<JsonElement> Allowed);

internal sealed record PpjSlotDefinitionModel(
    string Name,
    IReadOnlyList<string> Accepts,
    int Minimum,
    int Maximum);

internal sealed record PpjComponentBindingModel(string TargetId, string Field, string Parameter);

internal sealed record PpjComponentVariantModel(
    string Id,
    PpjConditionModel When,
    IReadOnlyList<PpjComponentAssignmentModel> Assignments);

internal sealed record PpjComponentAssignmentModel(
    string TargetId,
    string Field,
    JsonElement Value);

internal sealed record PpjConditionModel(
    string Argument,
    string Operator,
    JsonElement? Value);

internal sealed record PpjRepeatModel(
    IReadOnlyList<PpjRepeatItemModel> Items,
    string? Direction,
    double Gap);

internal sealed record PpjRepeatItemModel(
    string Key,
    IReadOnlyDictionary<string, JsonElement> Arguments);

internal sealed record PpjSectionModel(
    string Id,
    string Name,
    IReadOnlyList<string> PageIds,
    PpjNativeRefModel? NativeRef,
    JsonElement Raw);

internal sealed record PpjCustomShowModel(
    string Id,
    string Name,
    IReadOnlyList<string> PageIds,
    PpjNativeRefModel? NativeRef,
    JsonElement Raw);

internal sealed record PpjCommentModel(
    string Id,
    string PageId,
    string? TargetId,
    string? ParentId,
    string Author,
    string Text,
    bool Resolved,
    PpjNativeRefModel? NativeRef,
    JsonElement Raw);

internal static class PpjProgramParser
{
    internal static PpjProgramModel Parse(JsonElement root)
    {
        var meta = root.GetProperty("meta");
        var intent = root.GetProperty("intent");
        var brief = intent.GetProperty("brief");
        var design = root.GetProperty("design");
        var canvas = design.GetProperty("canvas");
        var styles = design.GetProperty("styles");

        return new PpjProgramModel(
            root.Clone(),
            new PpjMetaModel(
                meta.GetProperty("id").GetString()!,
                meta.GetProperty("title").GetString()!,
                meta.GetProperty("language").GetString()!,
                meta.GetProperty("version").GetInt32()),
            new PpjIntentModel(
                brief.GetProperty("primaryJob").GetString()!,
                brief.GetProperty("expectedOutcome").GetString()!,
                intent.GetProperty("audience").GetProperty("description").GetString()!,
                intent.GetProperty("narrative").GetProperty("thesis").GetString()!,
                intent.GetProperty("delivery").GetProperty("mode").GetString()!),
            new PpjDesignModel(
                canvas.GetProperty("width").GetDouble(),
                canvas.GetProperty("height").GetDouble(),
                canvas.TryGetProperty("nativeRef", out var canvasNativeRef) ? ParseNativeRef(canvasNativeRef) : null,
                IdSet(design.GetProperty("theme").GetProperty("colors")),
                IdSet(design.GetProperty("fonts")),
                OptionalIdSet(styles, "text"),
                OptionalIdSet(styles, "shape"),
                OptionalIdSet(styles, "chart"),
                OptionalIdSet(styles, "table"),
                design.GetProperty("motionPolicy").GetString()!,
                OptionalArray(design, "masters").Select(ParseMaster).ToArray(),
                OptionalArray(design, "layouts").Select(ParseLayout).ToArray()),
            OptionalArray(root, "assets").Select(ParseAsset).ToArray(),
            root.TryGetProperty("source", out var source) ? ParseSource(source) : null,
            OptionalArray(root, "components").Select(ParseComponent).ToArray(),
            root.GetProperty("pages").EnumerateArray().Select(ParsePage).ToArray(),
            OptionalArray(root, "sections").Select(item => new PpjSectionModel(
                item.GetProperty("id").GetString()!,
                item.GetProperty("name").GetString()!,
                Strings(item.GetProperty("pages")),
                item.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
                item.Clone())).ToArray(),
            OptionalArray(root, "customShows").Select(item => new PpjCustomShowModel(
                item.GetProperty("id").GetString()!,
                item.GetProperty("name").GetString()!,
                Strings(item.GetProperty("pages")),
                item.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
                item.Clone())).ToArray(),
            OptionalArray(root, "comments").Select(ParseComment).ToArray());
    }

    private static PpjAssetModel ParseAsset(JsonElement asset) => new(
        asset.GetProperty("id").GetString()!,
        asset.GetProperty("uri").GetString()!,
        asset.GetProperty("mimeType").GetString()!,
        asset.GetProperty("sha256").GetString()!,
        asset.GetProperty("rights").Clone(),
        asset.GetProperty("accessibility").Clone());

    private static PpjSourceModel ParseSource(JsonElement source)
    {
        var projection = source.GetProperty("projection");
        return new PpjSourceModel(
            source.GetProperty("kind").GetString()!,
            source.GetProperty("uri").GetString()!,
            source.GetProperty("sha256").GetString()!,
            source.GetProperty("revision").GetString()!,
            projection.GetProperty("sha256").GetString()!,
            projection.GetProperty("visibleObjectCount").GetInt32());
    }

    private static PpjMasterModel ParseMaster(JsonElement master)
    {
        var textStyles = master.TryGetProperty("textStyles", out var value) ? value : default;
        return new PpjMasterModel(
            master.GetProperty("id").GetString()!,
            master.GetProperty("name").GetString()!,
            master.TryGetProperty("background", out var background) ? background.Clone() : null,
            textStyles.ValueKind == JsonValueKind.Object ? OptionalArray(textStyles, "title").Select(item => item.Clone()).ToArray() : [],
            textStyles.ValueKind == JsonValueKind.Object ? OptionalArray(textStyles, "body").Select(item => item.Clone()).ToArray() : [],
            textStyles.ValueKind == JsonValueKind.Object ? OptionalArray(textStyles, "other").Select(item => item.Clone()).ToArray() : [],
            OptionalArray(master, "placeholders").Select(ParseLayoutPlaceholder).ToArray(),
            master.Clone());
    }

    private static PpjLayoutModel ParseLayout(JsonElement layout) => new(
        layout.GetProperty("id").GetString()!,
        layout.GetProperty("name").GetString()!,
        layout.GetProperty("master").GetString()!,
        layout.GetProperty("layoutType").GetString()!,
        layout.TryGetProperty("background", out var background) ? background.Clone() : null,
        OptionalArray(layout, "placeholders").Select(ParseLayoutPlaceholder).ToArray(),
        layout.Clone());

    private static PpjLayoutPlaceholderModel ParseLayoutPlaceholder(JsonElement placeholder) => new(
        placeholder.GetProperty("id").GetString()!,
        placeholder.GetProperty("name").GetString()!,
        placeholder.GetProperty("placeholderType").GetString()!,
        placeholder.GetProperty("index").GetUInt32(),
        ParseFrame(placeholder.GetProperty("frame")),
        placeholder.TryGetProperty("text", out var text) ? ParseText(text) : null,
        placeholder.TryGetProperty("style", out var style) ? style.Clone() : null,
        placeholder.Clone());

    private static PpjPageModel ParsePage(JsonElement page) => new(
        page.GetProperty("id").GetString()!,
        page.GetProperty("role").GetString()!,
        OptionalString(page, "name"),
        OptionalString(page, "claim"),
        OptionalString(page, "layout"),
        page.TryGetProperty("notes", out var notes) ? ParseText(notes) : null,
        page.TryGetProperty("hidden", out var hidden) ? hidden.GetBoolean() : null,
        page.GetProperty("elements").EnumerateArray().Select(ParseElement).ToArray(),
        OptionalArray(page, "animations").Select(ParseAnimation).ToArray(),
        page.TryGetProperty("transition", out var transition) ? ParseTransition(transition) : null,
        page.TryGetProperty("sourceClone", out var sourceClone) ? new PpjSourceCloneModel(
            sourceClone.GetProperty("page").GetString()!,
            sourceClone.GetProperty("capability").GetString()!,
            OptionalString(sourceClone, "retainElement")) : null,
        page.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
        page.Clone());

    private static PpjAnimationModel ParseAnimation(JsonElement animation) => new(
        animation.GetProperty("id").GetString()!,
        animation.GetProperty("target").GetString()!,
        animation.GetProperty("phase").GetString()!,
        animation.GetProperty("effect").GetString()!,
        OptionalString(animation, "direction"),
        animation.GetProperty("start").GetString()!,
        animation.GetProperty("durationMs").GetInt32(),
        OptionalInt(animation, "delayMs"),
        OptionalString(animation, "textBuild"),
        OptionalString(animation, "chartBuild"),
        OptionalInt(animation, "staggerMs"),
        animation.TryGetProperty("animateChartBackground", out var animateChartBackground)
            ? animateChartBackground.GetBoolean()
            : null);

    private static PpjTransitionModel ParseTransition(JsonElement transition) => new(
        transition.GetProperty("type").GetString()!,
        transition.TryGetProperty("durationMs", out var durationMs) ? durationMs.GetInt32() : null,
        OptionalString(transition, "direction"),
        OptionalString(transition, "orientation"),
        OptionalString(transition, "speed"),
        transition.TryGetProperty("throughBlack", out var throughBlack) ? throughBlack.GetBoolean() : null,
        transition.TryGetProperty("spokes", out var spokes) ? spokes.GetInt32() : null,
        transition.TryGetProperty("advanceOnClick", out var advanceOnClick) ? advanceOnClick.GetBoolean() : null,
        transition.TryGetProperty("advanceAfterMs", out var advanceAfterMs) ? advanceAfterMs.GetInt32() : null,
        OptionalString(transition, "fromPage"),
        OptionalArray(transition, "morphPairs").Select(pair => new PpjMorphPairModel(
            pair.GetProperty("key").GetString()!,
            pair.GetProperty("from").GetString()!,
            pair.GetProperty("to").GetString()!)).ToArray());

    private static PpjComponentModel ParseComponent(JsonElement component) => new(
        component.GetProperty("id").GetString()!,
        ParseFrame(component.GetProperty("frame")),
        OptionalArray(component, "parameters").Select(parameter => new PpjComponentParameterModel(
            parameter.GetProperty("name").GetString()!,
            parameter.GetProperty("type").GetString()!,
            OptionalBool(parameter, "required"),
            parameter.TryGetProperty("default", out var defaultValue) ? defaultValue.Clone() : null,
            OptionalArray(parameter, "allowed").Select(item => item.Clone()).ToArray())).ToArray(),
        OptionalArray(component, "slots").Select(slot => new PpjSlotDefinitionModel(
            slot.GetProperty("name").GetString()!,
            Strings(slot.GetProperty("accepts")),
            OptionalInt(slot, "minItems"),
            slot.TryGetProperty("maxItems", out var maximum) ? maximum.GetInt32() : 100000)).ToArray(),
        OptionalArray(component, "bindings").Select(binding => new PpjComponentBindingModel(
            binding.GetProperty("target").GetString()!,
            binding.GetProperty("field").GetString()!,
            binding.GetProperty("parameter").GetString()!)).ToArray(),
        OptionalArray(component, "variants").Select(variant => new PpjComponentVariantModel(
            variant.GetProperty("id").GetString()!,
            ParseCondition(variant.GetProperty("when")),
            variant.GetProperty("set").EnumerateArray().Select(assignment => new PpjComponentAssignmentModel(
                assignment.GetProperty("target").GetString()!,
                assignment.GetProperty("field").GetString()!,
                assignment.GetProperty("value").Clone())).ToArray())).ToArray(),
        component.GetProperty("elements").EnumerateArray().Select(ParseElement).ToArray(),
        component.Clone());

    private static PpjCommentModel ParseComment(JsonElement comment) => new(
        comment.GetProperty("id").GetString()!,
        comment.GetProperty("page").GetString()!,
        OptionalString(comment, "target"),
        OptionalString(comment, "parent"),
        comment.GetProperty("author").GetString()!,
        comment.GetProperty("text").GetString()!,
        comment.GetProperty("resolved").GetBoolean(),
        comment.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
        comment.Clone());

    internal static PpjElementModel ParseElement(JsonElement element)
    {
        var type = element.GetProperty("type").GetString()!;
        var common = new ElementCommon(
            element.GetProperty("id").GetString()!,
            type,
            ParseFrame(element.GetProperty("frame")),
            OptionalString(element, "name"),
            OptionalString(element, "role"),
            element.TryGetProperty("accessibility", out var accessibility) ? ParseAccessibility(accessibility) : null,
            element.TryGetProperty("hidden", out var hidden) ? hidden.GetBoolean() : null,
            element.TryGetProperty("locked", out var locked) ? locked.GetBoolean() : null,
            element.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
            element.Clone());

        PpjElementModel result = type switch
        {
            "text" => new PpjTextElementModel
            {
                Text = ParseText(element.GetProperty("text")),
                StyleRef = OptionalString(element, "styleRef"),
            },
            "shape" => new PpjShapeElementModel
            {
                GeometryKind = element.GetProperty("geometry").GetProperty("kind").GetString()!,
                GeometryPreset = OptionalString(element.GetProperty("geometry"), "preset"),
                GeometryAdjustments = ParsePresetAdjustments(element.GetProperty("geometry")),
                Text = element.TryGetProperty("text", out var text) ? ParseText(text) : null,
                StyleRef = OptionalString(element, "styleRef"),
            },
            "icon" => new PpjIconElementModel
            {
                IconName = element.GetProperty("iconName").GetString()!,
                StyleRef = OptionalString(element, "styleRef"),
            },
            "image" => new PpjImageElementModel
            {
                AssetId = element.GetProperty("asset").GetString()!,
                SvgAssetId = OptionalString(element, "svgAsset"),
                Fit = OptionalString(element, "fit"),
                MaskKind = element.TryGetProperty("mask", out var mask) ? OptionalString(mask, "kind") : null,
                MaskPreset = element.TryGetProperty("mask", out mask) ? OptionalString(mask, "preset") : null,
                MaskAdjustments = element.TryGetProperty("mask", out mask) ? ParsePresetAdjustments(mask) : [],
            },
            "chart" => new PpjChartElementModel
            {
                ChartType = element.GetProperty("chartType").GetString()!,
                Title = element.TryGetProperty("title", out var title) ? ParseText(title) : null,
                Data = ParseChartData(element.GetProperty("data")),
                StyleRef = OptionalString(element, "styleRef"),
            },
            "table" => new PpjTableElementModel
            {
                Columns = element.GetProperty("columns").EnumerateArray().Select(column => new PpjTableColumnModel(
                    OptionalString(column, "id"),
                    column.GetProperty("width").GetDouble())).ToArray(),
                Rows = element.GetProperty("rows").EnumerateArray().Select(ParseTableRow).ToArray(),
                StyleRef = OptionalString(element, "styleRef"),
            },
            "connector" => new PpjConnectorElementModel
            {
                ConnectorType = element.GetProperty("connectorType").GetString()!,
                From = ParseConnectorEndpoint(element.GetProperty("from")),
                To = ParseConnectorEndpoint(element.GetProperty("to")),
            },
            "group" => new PpjGroupElementModel
            {
                Elements = element.GetProperty("elements").EnumerateArray().Select(ParseElement).ToArray(),
            },
            "media" => new PpjMediaElementModel
            {
                MediaType = element.GetProperty("mediaType").GetString()!,
                AssetId = element.GetProperty("asset").GetString()!,
                PosterAssetId = element.GetProperty("posterAsset").GetString()!,
                StartAtMs = element.TryGetProperty("startAtMs", out var startAtMs) ? startAtMs.GetUInt64() : null,
                EndAtMs = element.TryGetProperty("endAtMs", out var endAtMs) ? endAtMs.GetUInt64() : null,
                Loop = element.TryGetProperty("loop", out var loop) ? loop.GetBoolean() : null,
                Mute = element.TryGetProperty("mute", out var mute) ? mute.GetBoolean() : null,
            },
            "placeholder" => new PpjPlaceholderElementModel
            {
                PlaceholderType = element.GetProperty("placeholderType").GetString()!,
                Text = element.TryGetProperty("text", out var placeholderText) ? ParseText(placeholderText) : null,
                StyleRef = OptionalString(element, "styleRef"),
            },
            "smartArt" => new PpjSmartArtElementModel
            {
                Mode = element.GetProperty("mode").GetString()!,
                Layout = OptionalString(element, "layout"),
                ShapeStyleRef = OptionalString(element, "shapeStyleRef"),
                TextStyleRef = OptionalString(element, "textStyleRef"),
                Nodes = element.GetProperty("nodes").EnumerateArray().Select(ParseSmartArtNode).ToArray(),
            },
            "ole" => new PpjOleElementModel
            {
                ProgramId = OptionalString(element, "programId"),
                PayloadAssetId = OptionalString(element, "payloadAsset"),
                PreviewAssetId = OptionalString(element, "previewAsset"),
            },
            "opaque" => new PpjOpaqueElementModel
            {
                NativeKind = element.GetProperty("nativeKind").GetString()!,
                Summary = element.GetProperty("summary").GetString()!,
                VisibleText = OptionalArray(element, "visibleText").Select(item => item.GetString()!).ToArray(),
                PreviewAssetId = OptionalString(element, "previewAsset"),
            },
            "component" => new PpjComponentElementModel
            {
                ComponentId = element.GetProperty("component").GetString()!,
                VariantId = OptionalString(element, "variant"),
                Arguments = ParseArguments(element, "arguments"),
                Slots = ParseSlots(element),
                Repeat = element.TryGetProperty("repeat", out var repeat) ? ParseRepeat(repeat) : null,
                When = element.TryGetProperty("when", out var when) ? ParseCondition(when) : null,
            },
            "slot" => new PpjSlotElementModel
            {
                SlotId = element.GetProperty("slot").GetString()!,
            },
            _ => throw new InvalidOperationException($"Validated PPJ element type {type} is not mapped."),
        };

        result.Id = common.Id;
        result.Type = common.Type;
        result.Frame = common.Frame;
        result.Name = common.Name;
        result.Role = common.Role;
        result.Accessibility = common.Accessibility;
        result.Hidden = common.Hidden;
        result.Locked = common.Locked;
        result.NativeRef = common.NativeRef;
        result.Raw = common.Raw;
        return result;
    }

    private static IReadOnlyList<int> ParsePresetAdjustments(JsonElement geometry) =>
        geometry.GetProperty("kind").GetString() == "preset" &&
        geometry.TryGetProperty("adjustments", out var adjustments)
            ? adjustments.EnumerateArray().Select(value => value.GetInt32()).ToArray()
            : [];

    private static PpjChartDataModel ParseChartData(JsonElement data) => new(
        data.GetProperty("categories").EnumerateArray().Select(item => item.Clone()).ToArray(),
        data.GetProperty("series").EnumerateArray().Select(series => new PpjChartSeriesModel(
            series.GetProperty("id").GetString()!,
            series.GetProperty("name").GetString()!,
            series.GetProperty("values").EnumerateArray().Select(value => value.ValueKind == JsonValueKind.Null ? (double?)null : value.GetDouble()).ToArray(),
            series.TryGetProperty("pointRoles", out var pointRoles)
                ? pointRoles.EnumerateArray().Select(value => value.GetString()!).ToArray()
                : [],
            series.TryGetProperty("xValues", out var xValues)
                ? xValues.EnumerateArray().Select(value => value.GetDouble()).ToArray()
                : [],
            series.TryGetProperty("bubbleSizes", out var bubbleSizes)
                ? bubbleSizes.EnumerateArray().Select(value => value.GetDouble()).ToArray()
                : [],
            series.TryGetProperty("openValues", out var openValues)
                ? openValues.EnumerateArray().Select(value => value.GetDouble()).ToArray()
                : [],
            series.TryGetProperty("highValues", out var highValues)
                ? highValues.EnumerateArray().Select(value => value.GetDouble()).ToArray()
                : [],
            series.TryGetProperty("lowValues", out var lowValues)
                ? lowValues.EnumerateArray().Select(value => value.GetDouble()).ToArray()
                : [],
            series.TryGetProperty("parents", out var parents)
                ? parents.EnumerateArray().Select(value => value.ValueKind == JsonValueKind.Null ? null : value.GetString()).ToArray()
                : [],
            series.TryGetProperty("sources", out var sources)
                ? sources.EnumerateArray().Select(value => value.GetString()!).ToArray()
                : [],
            series.TryGetProperty("targets", out var targets)
                ? targets.EnumerateArray().Select(value => value.GetString()!).ToArray()
                : [],
            series.TryGetProperty("levels", out var levels) ? levels.GetInt32() : null,
            OptionalString(series, "chartType"),
            OptionalString(series, "axis"),
            series.Clone())).ToArray());

    private static PpjTableRowModel ParseTableRow(JsonElement row) => new(
        OptionalString(row, "id"),
        row.TryGetProperty("height", out var height) ? height.GetDouble() : null,
        row.GetProperty("cells").EnumerateArray().Select(cell => new PpjTableCellModel(
            OptionalString(cell, "id"),
            ParseText(cell.GetProperty("text")),
            cell.TryGetProperty("rowSpan", out var rowSpan) ? rowSpan.GetInt32() : 1,
            cell.TryGetProperty("columnSpan", out var columnSpan) ? columnSpan.GetInt32() : 1)).ToArray());

    private static PpjConnectorEndpointModel ParseConnectorEndpoint(JsonElement endpoint) => new(
        OptionalString(endpoint, "element"),
        OptionalString(endpoint, "anchor"),
        endpoint.TryGetProperty("x", out var x) ? x.GetDouble() : null,
        endpoint.TryGetProperty("y", out var y) ? y.GetDouble() : null);

    private static PpjSmartArtNodeModel ParseSmartArtNode(JsonElement node) => new(
        node.GetProperty("id").GetString()!,
        OptionalString(node, "parent"),
        ParseText(node.GetProperty("text")),
        OptionalString(node, "styleRef"),
        OptionalString(node, "shapeStyleRef"),
        OptionalString(node, "asset"),
        node.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
        node.Clone());

    private static PpjTextContentModel ParseText(JsonElement text)
    {
        if (text.ValueKind == JsonValueKind.String)
            return new PpjTextContentModel(text.GetString(), []);
        return new PpjTextContentModel(
            null,
            text.GetProperty("paragraphs").EnumerateArray().Select(paragraph => new PpjParagraphModel(
                OptionalString(paragraph, "id"),
                paragraph.GetProperty("runs").EnumerateArray().Select(run => new PpjRunModel(
                    OptionalString(run, "id"),
                    OptionalString(run, "text"),
                    run.TryGetProperty("formula", out var formula)
                        ? new PpjFormulaModel(formula.GetProperty("syntax").GetString()!, formula.GetProperty("source").GetString()!)
                        : null)).ToArray())).ToArray());
    }

    private static PpjFrameModel ParseFrame(JsonElement frame) => new(
        frame.GetProperty("x").GetDouble(),
        frame.GetProperty("y").GetDouble(),
        frame.GetProperty("width").GetDouble(),
        frame.GetProperty("height").GetDouble(),
        frame.TryGetProperty("rotation", out var rotation) ? rotation.GetDouble() : 0,
        OptionalBool(frame, "flipH"),
        OptionalBool(frame, "flipV"));

    private static PpjAccessibilityModel ParseAccessibility(JsonElement accessibility) => new(
        accessibility.GetProperty("decorative").GetBoolean(),
        OptionalString(accessibility, "title"),
        OptionalString(accessibility, "description"));

    private static PpjNativeRefModel ParseNativeRef(JsonElement nativeRef) => new(
        nativeRef.GetProperty("handle").GetString()!,
        nativeRef.GetProperty("sourceSha256").GetString()!,
        nativeRef.GetProperty("revision").GetString()!,
        nativeRef.GetProperty("objectHash").GetString()!,
        nativeRef.GetProperty("capabilitySetSha256").GetString()!,
        nativeRef.GetProperty("capabilities").EnumerateArray().Select(capability => new PpjNativeCapabilityModel(
            capability.GetProperty("id").GetString()!,
            capability.GetProperty("operation").GetString()!,
            capability.GetProperty("expectedHash").GetString()!,
            Strings(capability.GetProperty("fields")))).ToArray(),
        nativeRef.TryGetProperty("leaves", out var leaves)
            ? leaves.EnumerateArray().Select(leaf => new PpjNativeLeafModel(
                leaf.GetProperty("id").GetString()!,
                leaf.GetProperty("kind").GetString()!,
                leaf.GetProperty("expectedHash").GetString()!,
                leaf.GetProperty("value").Clone())).ToArray()
            : []);

    private static PpjRepeatModel ParseRepeat(JsonElement repeat)
    {
        var layout = repeat.TryGetProperty("layout", out var layoutElement) ? layoutElement : default;
        return new PpjRepeatModel(
            repeat.GetProperty("items").EnumerateArray().Select(item => new PpjRepeatItemModel(
                item.GetProperty("key").GetString()!,
                ParseArguments(item, "arguments"))).ToArray(),
            layout.ValueKind == JsonValueKind.Object ? OptionalString(layout, "direction") : null,
            layout.ValueKind == JsonValueKind.Object && layout.TryGetProperty("gap", out var gap) ? gap.GetDouble() : 0);
    }

    private static PpjConditionModel ParseCondition(JsonElement condition) => new(
        condition.GetProperty("argument").GetString()!,
        condition.GetProperty("operator").GetString()!,
        condition.TryGetProperty("value", out var value) ? value.Clone() : null);

    private static IReadOnlyDictionary<string, JsonElement> ParseArguments(JsonElement owner, string name)
    {
        if (!owner.TryGetProperty(name, out var arguments))
            return new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        return arguments.EnumerateObject().ToDictionary(
            property => property.Name,
            property => property.Value.Clone(),
            StringComparer.Ordinal);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<PpjElementModel>> ParseSlots(JsonElement element)
    {
        if (!element.TryGetProperty("slots", out var slots))
            return new Dictionary<string, IReadOnlyList<PpjElementModel>>(StringComparer.Ordinal);
        return slots.EnumerateObject().ToDictionary(
            property => property.Name,
            property => (IReadOnlyList<PpjElementModel>)property.Value.EnumerateArray().Select(ParseElement).ToArray(),
            StringComparer.Ordinal);
    }

    private static IReadOnlySet<string> IdSet(JsonElement array) =>
        array.EnumerateArray().Select(item => item.GetProperty("id").GetString()!).ToHashSet(StringComparer.Ordinal);

    private static IReadOnlySet<string> OptionalIdSet(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var array)
            ? IdSet(array)
            : new HashSet<string>(StringComparer.Ordinal);

    private static IReadOnlyList<string> Strings(JsonElement array) =>
        array.EnumerateArray().Select(item => item.GetString()!).ToArray();

    private static IEnumerable<JsonElement> OptionalArray(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var array) ? array.EnumerateArray().Select(item => item.Clone()) : [];

    private static string? OptionalString(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) ? value.GetString() : null;

    private static bool OptionalBool(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) && value.GetBoolean();

    private static int OptionalInt(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) ? value.GetInt32() : 0;

    private sealed record ElementCommon(
        string Id,
        string Type,
        PpjFrameModel Frame,
        string? Name,
        string? Role,
        PpjAccessibilityModel? Accessibility,
        bool? Hidden,
        bool? Locked,
        PpjNativeRefModel? NativeRef,
        JsonElement Raw);
}
