using System.Text.Json;
using System.Text.Json.Nodes;

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
    IReadOnlySet<string> ImageStyleIds,
    IReadOnlySet<string> ChartStyleIds,
    IReadOnlySet<string> TableStyleIds,
    string MotionPolicy,
    IReadOnlyList<PpjMasterModel> Masters,
    IReadOnlyList<PpjLayoutModel> Layouts);

internal sealed record PpjMasterModel(
    string Id,
    string Name,
    PpjNativeRefModel? NativeRef,
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
    PpjNativeRefModel? NativeRef,
    JsonElement? Background,
    IReadOnlyList<PpjLayoutPlaceholderModel> Placeholders,
    JsonElement Raw);

internal sealed record PpjLayoutPlaceholderModel(
    string Id,
    string Name,
    string PlaceholderType,
    uint Index,
    PpjFrameModel Frame,
    PpjNativeRefModel? NativeRef,
    PpjTextContentModel? Text,
    JsonElement? Style,
    JsonElement Raw);

internal sealed record PpjAssetModel(
    string Id,
    string Uri,
    string MimeType,
    string Sha256,
    JsonElement Rights,
    JsonElement Accessibility,
    int? WidthPx,
    int? HeightPx);

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
    internal string? StyleRef { get; init; }
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
    string CategoryFormula,
    string XValueFormula,
    string ValueFormula,
    string BubbleSizeFormula,
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
    int? XAxisIndex,
    int? YAxisIndex,
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
    int ColumnSpan,
    JsonElement Raw);

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
    internal PpjFrameModel ChildFrame { get; init; } = new(0, 0, 1, 1, 0, false, false);
    internal IReadOnlyList<string> ReadingOrder { get; init; } = [];
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
    internal string? PlaybackTrigger { get; init; }
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
    internal string? DefinitionAssetId { get; init; }
    internal string? LayoutDefinitionId { get; init; }
    internal bool DetachToShapes { get; init; }
    internal string? ShapeStyleRef { get; init; }
    internal string? TextStyleRef { get; init; }
    internal required IReadOnlyList<PpjSmartArtNodeModel> Nodes { get; init; }
    internal required IReadOnlyList<PpjSmartArtConnectionModel> Connections { get; init; }
}

internal sealed record PpjSmartArtNodeModel(
    string Id,
    PpjTextContentModel Text,
    string? StyleRef,
    string? ShapeStyleRef,
    string? AssetId,
    JsonElement? Image,
    PpjNativeRefModel? NativeRef,
    JsonElement Raw);

internal sealed record PpjSmartArtConnectionModel(
    string Id,
    string FromId,
    string ToId,
    string Role,
    uint Order);

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

internal sealed record PpjTextFieldModel(string? Id, string Type, string Text);

internal sealed record PpjRunModel(string? Id, string? Text, PpjFormulaModel? Formula, PpjTextFieldModel? Field, bool LineBreak);

internal sealed record PpjPageModel(
    string Id,
    string Role,
    string? Name,
    string? Claim,
    string? LayoutId,
    PpjTextContentModel? Notes,
    bool? Hidden,
    IReadOnlyList<PpjElementModel> Elements,
    IReadOnlyList<string> ReadingOrder,
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
    bool? AnimateChartBackground,
    string? Easing = null,
    int? Repeat = null,
    bool? AutoReverse = null);

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
    int Maximum,
    PpjImageSlotPolicyModel? ImagePolicy);

internal sealed record PpjImageSlotPolicyModel(
    string Role,
    IReadOnlyList<string> AllowedFit,
    IReadOnlyList<string> AllowedMask,
    int? MinimumWidthPx,
    int? MinimumHeightPx,
    IReadOnlySet<string> Rights);

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
    double Gap,
    int? Columns,
    double? RowGap,
    string? Anchor,
    IReadOnlyList<double>? Weights);

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
    string Kind,
    string Author,
    string Text,
    bool Resolved,
    string? Status,
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
                OptionalIdSet(styles, "image"),
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
        asset.GetProperty("accessibility").Clone(),
        asset.TryGetProperty("widthPx", out var width) ? width.GetInt32() : null,
        asset.TryGetProperty("heightPx", out var height) ? height.GetInt32() : null);

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
            master.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
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
        layout.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
        layout.TryGetProperty("background", out var background) ? background.Clone() : null,
        OptionalArray(layout, "placeholders").Select(ParseLayoutPlaceholder).ToArray(),
        layout.Clone());

    private static PpjLayoutPlaceholderModel ParseLayoutPlaceholder(JsonElement placeholder) => new(
        placeholder.GetProperty("id").GetString()!,
        placeholder.GetProperty("name").GetString()!,
        placeholder.GetProperty("placeholderType").GetString()!,
        placeholder.GetProperty("index").GetUInt32(),
        ParseFrame(placeholder.GetProperty("frame")),
        placeholder.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
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
        OptionalArray(page, "readingOrder").Select(item => item.GetString()!).ToArray(),
        OptionalArray(page, "animations").Select(ParseAnimation).Concat(ParseTimingAnimations(page)).ToArray(),
        page.TryGetProperty("transition", out var transition) ? ParseTransition(transition) : null,
        page.TryGetProperty("sourceClone", out var sourceClone) ? new PpjSourceCloneModel(
            sourceClone.GetProperty("page").GetString()!,
            sourceClone.GetProperty("capability").GetString()!,
            OptionalString(sourceClone, "retainElement")) : null,
        page.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
        page.Clone());

    private static PpjAnimationModel ParseAnimation(JsonElement animation)
    {
        // `trigger` is the explicit animation-array form of the timing graph
        // sugar.  `timeline` leaves the required start condition authoritative;
        // every other value is normalized into the same wire start condition.
        var start = animation.GetProperty("start").GetString()!;
        var trigger = OptionalString(animation, "trigger");
        return new(
            animation.GetProperty("id").GetString()!,
            animation.GetProperty("target").GetString()!,
            animation.GetProperty("phase").GetString()!,
            animation.GetProperty("effect").GetString()!,
            OptionalString(animation, "direction"),
            trigger is null or "timeline" ? start : trigger,
            animation.GetProperty("durationMs").GetInt32(),
            OptionalInt(animation, "delayMs"),
            OptionalString(animation, "textBuild"),
            OptionalString(animation, "chartBuild"),
            OptionalInt(animation, "staggerMs"),
            animation.TryGetProperty("animateChartBackground", out var animateChartBackground)
                ? animateChartBackground.GetBoolean()
                : null,
            OptionalString(animation, "easing"),
            animation.TryGetProperty("repeat", out var repeat) ? repeat.GetInt32() : null,
            animation.TryGetProperty("autoReverse", out var autoReverse) ? autoReverse.GetBoolean() : null);
    }

    private static IEnumerable<PpjAnimationModel> ParseTimingAnimations(JsonElement page)
    {
        if (!page.TryGetProperty("timing", out var timing) ||
            !timing.TryGetProperty("nodes", out var nodes))
            yield break;
        foreach (var node in nodes.EnumerateArray())
        {
            // `trigger` is authored sugar for the existing PresentationML
            // start condition.  `timeline` means that the required `start`
            // field remains authoritative; the other bounded trigger names
            // are normalized to the same wire value so the graph cannot
            // silently degrade into a default fade.
            var start = node.GetProperty("start").GetString()!;
            var trigger = OptionalString(node, "trigger");
            yield return new PpjAnimationModel(
                node.GetProperty("id").GetString()!,
                node.GetProperty("target").GetString()!,
                node.GetProperty("phase").GetString()!,
                node.GetProperty("effect").GetString()!,
                OptionalString(node, "direction"),
                trigger is null or "timeline" ? start : trigger,
                node.GetProperty("durationMs").GetInt32(),
                OptionalInt(node, "delayMs"),
                OptionalString(node, "textBuild"),
                OptionalString(node, "chartBuild"),
                OptionalInt(node, "staggerMs"),
                node.TryGetProperty("animateChartBackground", out var animateChartBackground)
                    ? animateChartBackground.GetBoolean()
                    : null,
                OptionalString(node, "easing"),
                node.TryGetProperty("repeat", out var repeat) ? repeat.GetInt32() : null,
                node.TryGetProperty("autoReverse", out var autoReverse) ? autoReverse.GetBoolean() : null);
        }
    }

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
            slot.TryGetProperty("maxItems", out var maximum) ? maximum.GetInt32() : 100000,
            slot.TryGetProperty("imagePolicy", out var imagePolicy) ? new PpjImageSlotPolicyModel(
                imagePolicy.GetProperty("role").GetString()!,
                OptionalArray(imagePolicy, "allowedFit").Select(item => item.GetString()!).ToArray(),
                OptionalArray(imagePolicy, "allowedMask").Select(item => item.GetString()!).ToArray(),
                imagePolicy.TryGetProperty("minWidthPx", out var minWidth) ? minWidth.GetInt32() : null,
                imagePolicy.TryGetProperty("minHeightPx", out var minHeight) ? minHeight.GetInt32() : null,
                OptionalArray(imagePolicy, "rights").Select(item => item.GetString()!).ToHashSet(StringComparer.Ordinal))
                : null)).ToArray(),
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
        OptionalString(comment, "kind") ?? "legacy",
        comment.GetProperty("author").GetString()!,
        comment.GetProperty("text").GetString()!,
        comment.GetProperty("resolved").GetBoolean(),
        OptionalString(comment, "status"),
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
            "line" => new PpjShapeElementModel
            {
                GeometryKind = "line",
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
                StyleRef = OptionalString(element, "styleRef"),
                // `fit` may be a typed design-grammar reference.  Keep the
                // model's convenience value literal-only; compiler paths
                // resolve the raw token against the program catalog so a
                // token object never reaches JsonElement.GetString().
                Fit = OptionalLiteralString(element, "fit"),
                MaskKind = element.TryGetProperty("mask", out var mask) ? OptionalString(mask, "kind") : null,
                MaskPreset = element.TryGetProperty("mask", out mask) ? OptionalString(mask, "preset") : null,
                MaskAdjustments = element.TryGetProperty("mask", out mask) ? ParsePresetAdjustments(mask) : [],
            },
            "chart" => new PpjChartElementModel
            {
                ChartType = ChartTypeForElement(element),
                Title = element.TryGetProperty("title", out var title) ? ParseText(title) : null,
                Data = ParseChartData(element.GetProperty("data"), ChartTypeForElement(element)),
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
                ChildFrame = element.TryGetProperty("childFrame", out var childFrame)
                    ? ParseFrame(childFrame)
                    : common.Frame,
                ReadingOrder = OptionalArray(element, "readingOrder").Select(item => item.GetString()!).ToArray(),
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
                PlaybackTrigger = element.TryGetProperty("playback", out var playback)
                    ? playback.GetProperty("trigger").GetString()
                    : null,
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
                DefinitionAssetId = OptionalString(element, "definitionAsset"),
                LayoutDefinitionId = OptionalString(element, "layoutDefinitionId"),
                DetachToShapes = element.TryGetProperty("detachToShapes", out var detachToShapes) && detachToShapes.GetBoolean(),
                ShapeStyleRef = OptionalString(element, "shapeStyleRef"),
                TextStyleRef = OptionalString(element, "textStyleRef"),
                Nodes = element.GetProperty("nodes").EnumerateArray().Select(ParseSmartArtNode).ToArray(),
                Connections = element.TryGetProperty("connections", out var connections)
                    ? connections.EnumerateArray().Select(ParseSmartArtConnection).ToArray()
                    : [],
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
        if (result is PpjChartElementModel &&
            (common.Raw.TryGetProperty("data", out var chartData) && chartData.TryGetProperty("dataset", out _) ||
             common.Raw.TryGetProperty("xAxis", out var rawXAxis) && rawXAxis.ValueKind == JsonValueKind.Array ||
             common.Raw.TryGetProperty("yAxis", out var rawYAxis) && rawYAxis.ValueKind == JsonValueKind.Array))
        {
            var canonical = JsonNode.Parse(common.Raw.GetRawText())!.AsObject();
            NormalizeChartAxisArray(canonical, "xAxis", "secondaryXAxis");
            NormalizeChartAxisArray(canonical, "yAxis", "secondaryYAxis");
            if (common.Raw.TryGetProperty("data", out chartData) && chartData.TryGetProperty("dataset", out _))
            {
                var chartType = ChartTypeForElement(element);
                var canonicalData = chartData.TryGetProperty("series", out var datasetSeries) &&
                    datasetSeries.ValueKind == JsonValueKind.Array &&
                    datasetSeries.GetArrayLength() > 0 &&
                    datasetSeries[0].TryGetProperty("encode", out _)
                    ? CanonicalDatasetSeries(chartData, chartType)
                    : CanonicalDataset(chartData, chartType);
                canonical["data"] = JsonNode.Parse(canonicalData.GetRawText());
            }
            result.Raw = JsonDocument.Parse(canonical.ToJsonString()).RootElement.Clone();
        }
        return result;
    }

    private static void NormalizeChartAxisArray(JsonObject chart, string primaryName, string secondaryName)
    {
        if (chart[primaryName] is not JsonArray axes) return;
        if (axes.Count is < 1 or > 2 || axes.Any(axis => axis is null))
            throw new InvalidOperationException($"{primaryName} arrays are bounded to one or two axis objects.");
        if (chart.ContainsKey(secondaryName))
            throw new InvalidOperationException($"{primaryName} array cannot be combined with {secondaryName}.");
        chart[primaryName] = axes[0]!.DeepClone();
        if (axes.Count == 2) chart[secondaryName] = axes[1]!.DeepClone();
    }

    private static IReadOnlyList<int> ParsePresetAdjustments(JsonElement geometry) =>
        geometry.GetProperty("kind").GetString() == "preset" &&
        geometry.TryGetProperty("adjustments", out var adjustments)
            ? adjustments.EnumerateArray().Select(value => value.GetInt32()).ToArray()
            : [];

    private static PpjChartDataModel ParseChartData(JsonElement data, string chartType)
    {
        if (!data.TryGetProperty("dataset", out _)) return ParseLegacyChartData(data);
        if (data.TryGetProperty("series", out var series) &&
            series.ValueKind == JsonValueKind.Array &&
            series.GetArrayLength() > 0 &&
            series[0].TryGetProperty("encode", out _))
            return ParseLegacyChartData(CanonicalDatasetSeries(data, chartType));
        return ParseLegacyChartData(CanonicalDataset(data, chartType));
    }

    private static string ChartTypeForElement(JsonElement element)
    {
        if (element.TryGetProperty("chartType", out var chartType)) return chartType.GetString()!;
        var data = element.GetProperty("data");
        if (data.TryGetProperty("series", out var series) && series.GetArrayLength() > 0)
        {
            var types = series.EnumerateArray()
                .Select(item => item.TryGetProperty("type", out var type) ? type.GetString() : item.TryGetProperty("chartType", out var chartTypeValue) ? chartTypeValue.GetString() : null)
                .Where(type => !string.IsNullOrEmpty(type))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            if (types.Contains("candlestick", StringComparer.Ordinal) &&
                types.All(type => type is "candlestick" or "line" or "area" or "bar"))
                return "candlestick";
            if (types.Length == 1) return types[0] == "bar" ? "bar" : types[0]!;
            if (types.Length > 1) return "combo";
        }
        throw new InvalidOperationException("Chart requires chartType or a dataset series type.");
    }

    private static PpjChartDataModel ParseLegacyChartData(JsonElement data) => new(
        data.GetProperty("categories").EnumerateArray().Select(item => item.Clone()).ToArray(),
        data.GetProperty("series").EnumerateArray().Select(series => new PpjChartSeriesModel(
            series.GetProperty("id").GetString()!,
            series.GetProperty("name").GetString()!,
            OptionalString(series, "categoryFormula") ?? string.Empty,
            OptionalString(series, "xValueFormula") ?? string.Empty,
            OptionalString(series, "valueFormula") ?? string.Empty,
            OptionalString(series, "bubbleSizeFormula") ?? string.Empty,
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
            series.TryGetProperty("xAxisIndex", out var xAxisIndex) ? xAxisIndex.GetInt32() : null,
            series.TryGetProperty("yAxisIndex", out var yAxisIndex) ? yAxisIndex.GetInt32() : null,
            series.Clone())).ToArray());

    private static JsonElement CanonicalDataset(JsonElement data, string chartType)
    {
        var dataset = data.GetProperty("dataset");
        var columnNames = dataset.GetProperty("cols").EnumerateArray().Select((column, index) =>
            column.ValueKind == JsonValueKind.String
                ? column.GetString()!
                : column.TryGetProperty("id", out var id) ? id.GetString()! : $"col-{index + 1}").ToArray();
        var rows = dataset.GetProperty("rows").EnumerateArray().Select(row => DatasetRow(row, columnNames)).ToArray();
        if (data.TryGetProperty("dataFilter", out var filters))
            rows = rows.Where(row => filters.EnumerateArray().All(filter => MatchesFilter(row, filter, columnNames))).ToArray();

        var encoding = data.GetProperty("encoding");
        var categoryColumn = RefIndex(encoding, "category", columnNames);
        var xColumn = RefIndex(encoding, "x", columnNames) ?? categoryColumn;
        var valueColumn = RefIndex(encoding, "value", columnNames) ?? RefIndex(encoding, "y", columnNames);
        var seriesColumn = RefIndex(encoding, "series", columnNames);
        var openColumn = RefIndex(encoding, "open", columnNames);
        var highColumn = RefIndex(encoding, "high", columnNames);
        var lowColumn = RefIndex(encoding, "low", columnNames);
        var closeColumn = RefIndex(encoding, "close", columnNames);
        var sizeColumn = RefIndex(encoding, "size", columnNames);
        var parentColumn = RefIndex(encoding, "parent", columnNames);
        var sourceColumn = RefIndex(encoding, "source", columnNames);
        var targetColumn = RefIndex(encoding, "target", columnNames);
        var flowColumn = RefIndex(encoding, "flow", columnNames);
        var isTotalColumn = RefIndex(encoding, "isTotal", columnNames);
        var levelColumn = RefIndex(encoding, "level", columnNames);
        valueColumn ??= chartType == "candlestick" ? closeColumn : null;
        valueColumn ??= chartType == "sankey" ? flowColumn : null;
        if (valueColumn is null)
            throw new InvalidOperationException("Chart dataset encoding requires value or y.");

        var numericX = chartType is "scatter" or "bubble";
        var categoryValues = numericX
            ? Array.Empty<JsonElement>()
            : rows.Select((row, index) => categoryColumn is { } column ? row[column] : JsonDocument.Parse($"\"{index + 1}\"").RootElement.Clone())
                .GroupBy(ScalarKey, StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
        var grouped = rows.GroupBy(row => seriesColumn is { } column ? ScalarKey(row[column]) : "Series 1", StringComparer.Ordinal).ToArray();
        var series = new JsonArray();
        foreach (var (group, groupIndex) in grouped.Select((group, index) => (group, index)))
        {
            var sourceRows = group.ToArray();
            var item = new JsonObject
            {
                ["id"] = $"series-{groupIndex + 1}",
                ["name"] = group.Key,
                ["values"] = new JsonArray(),
            };
            var values = (JsonArray)item["values"]!;
            if (numericX)
            {
                item["xValues"] = new JsonArray(sourceRows.Select(row => NumberNode(row[xColumn!.Value])).ToArray());
                foreach (var row in sourceRows) values.Add(NumberNode(row[valueColumn.Value]));
                if (chartType == "bubble" && sizeColumn is { } bubbleColumn)
                    item["bubbleSizes"] = new JsonArray(sourceRows.Select(row => NumberNode(row[bubbleColumn])).ToArray());
            }
            else
            {
                foreach (var category in categoryValues)
                {
                    var row = categoryColumn is { } categoryIndex
                        ? sourceRows.FirstOrDefault(candidate => ScalarKey(candidate[categoryIndex]) == ScalarKey(category))
                        : sourceRows.ElementAtOrDefault(Array.IndexOf(categoryValues, category));
                    values.Add(row is null ? null : NumberNode(row[valueColumn.Value]));
                }
            }
            AddEncodedArray(item, "openValues", sourceRows, openColumn, numericX, categoryValues, categoryColumn, NumberNode);
            AddEncodedArray(item, "highValues", sourceRows, highColumn, numericX, categoryValues, categoryColumn, NumberNode);
            AddEncodedArray(item, "lowValues", sourceRows, lowColumn, numericX, categoryValues, categoryColumn, NumberNode);
            AddEncodedArray(item, "parents", sourceRows, parentColumn, numericX, categoryValues, categoryColumn, ValueNode);
            AddEncodedArray(item, "sources", sourceRows, sourceColumn, numericX, categoryValues, categoryColumn, ValueNode);
            AddEncodedArray(item, "targets", sourceRows, targetColumn, numericX, categoryValues, categoryColumn, ValueNode);
            if (isTotalColumn is { } totalColumn)
                item["pointRoles"] = new JsonArray(sourceRows.Select(row =>
                    (JsonNode?)JsonValue.Create(row[totalColumn].ValueKind == JsonValueKind.True ? "total" : "delta")).ToArray());
            if (levelColumn is { } level)
            {
                var levelValues = sourceRows.Select(row => row[level]).Where(value => value.ValueKind != JsonValueKind.Null).Select(value => value.GetInt32()).Distinct().ToArray();
                if (levelValues.Length == 1) item["levels"] = levelValues[0];
            }
            if (data.TryGetProperty("seriesDefaults", out var defaults))
                ApplySeriesDefaults(item, defaults, chartType, chartType);
            series.Add(item);
        }
        var output = new JsonObject
        {
            ["categories"] = new JsonArray(categoryValues.Select(ToNode).ToArray()),
            ["series"] = series,
        };
        return JsonDocument.Parse(output.ToJsonString()).RootElement.Clone();
    }

    private static JsonElement CanonicalDatasetSeries(JsonElement data, string chartType)
    {
        var dataset = data.GetProperty("dataset");
        var columnNames = dataset.GetProperty("cols").EnumerateArray().Select((column, index) =>
            column.ValueKind == JsonValueKind.String
                ? column.GetString()!
                : column.TryGetProperty("id", out var id) ? id.GetString()! : $"col-{index + 1}").ToArray();
        IReadOnlyList<JsonElement[]> allRows = dataset.GetProperty("rows").EnumerateArray().Select(row => DatasetRow(row, columnNames)).ToArray();
        allRows = ApplyFilters(allRows, data.TryGetProperty("dataFilter", out var filters) ? filters : default, columnNames);
        var definitions = data.GetProperty("series").EnumerateArray().ToArray();
        var definitionTypes = definitions.Select(item => item.GetProperty("type").GetString()!).ToArray();
        if (definitionTypes.Any(type => type == "heatmap"))
        {
            if (definitionTypes.Any(type => type != "heatmap"))
                throw new InvalidOperationException("Heatmap series cannot be mixed with other dataset series types.");
            return CanonicalHeatmapDatasetSeries(data, columnNames, allRows, definitions);
        }
        if (definitionTypes.Any(type => type == "sankey") && definitionTypes.Any(type => type != "sankey"))
            throw new InvalidOperationException("Sankey series cannot be mixed with other dataset series types.");

        var outputSeries = new JsonArray();
        JsonElement[] categories;
        if (definitionTypes.All(type => type is "scatter" or "bubble"))
            categories = [];
        else if (definitionTypes.Any(type => type == "sankey"))
        {
            var sankey = definitions.Single(item => item.GetProperty("type").GetString() == "sankey");
            var encode = sankey.GetProperty("encode");
            var source = RefIndex(encode, "source", columnNames) ?? throw new InvalidOperationException("Sankey encoding requires source.");
            var target = RefIndex(encode, "target", columnNames) ?? throw new InvalidOperationException("Sankey encoding requires target.");
            categories = allRows.SelectMany(row => new[] { row[source], row[target] }).GroupBy(ScalarKey, StringComparer.Ordinal).Select(group => group.First()).ToArray();
        }
        else
        {
            var first = definitions[0].GetProperty("encode");
            var categoryColumn = RefIndex(first, "category", columnNames) ?? RefIndex(first, "x", columnNames);
            categories = allRows.Select((row, index) => categoryColumn is { } column ? row[column] : JsonDocument.Parse($"\"{index + 1}\"").RootElement.Clone())
                .GroupBy(ScalarKey, StringComparer.Ordinal).Select(group => group.First()).ToArray();
        }

        foreach (var (definition, definitionIndex) in definitions.Select((item, index) => (item, index)))
        {
            var type = definition.GetProperty("type").GetString()!;
            var xAxisIndex = definition.TryGetProperty("xAxisIndex", out var xAxisIndexValue) ? xAxisIndexValue.GetInt32() : 0;
            var yAxisIndex = definition.TryGetProperty("yAxisIndex", out var yAxisIndexValue) ? yAxisIndexValue.GetInt32() : 0;
            if (xAxisIndex > 1 || yAxisIndex > 1)
                throw new InvalidOperationException("Dataset series axis indexes are bounded to primary (0) and secondary (1) axes.");
            var indexedSecondary = xAxisIndex == 1 || yAxisIndex == 1;
            if (indexedSecondary && chartType != "combo")
                throw new InvalidOperationException("Dataset series secondary axis indexes require a combo chart.");
            if (indexedSecondary && definition.TryGetProperty("axis", out var explicitAxis) && explicitAxis.GetString() == "primary")
                throw new InvalidOperationException("Dataset series axis=primary conflicts with a secondary axis index.");
            var encode = definition.GetProperty("encode");
            var rows = ApplyFilters(allRows, definition.TryGetProperty("dataFilter", out var seriesFilters) ? seriesFilters : default, columnNames);
            var categoryColumn = RefIndex(encode, "category", columnNames) ?? RefIndex(encode, "x", columnNames);
            var xColumn = RefIndex(encode, "x", columnNames) ?? categoryColumn;
            var valueColumn = RefIndex(encode, "value", columnNames) ?? RefIndex(encode, "y", columnNames);
            valueColumn ??= type == "candlestick" ? RefIndex(encode, "close", columnNames) : null;
            valueColumn ??= type == "sankey" ? RefIndex(encode, "flow", columnNames) : null;
            if (valueColumn is null && type != "sankey")
                throw new InvalidOperationException($"Dataset series {definitionIndex + 1} requires a value, y, or type-specific value channel.");
            var item = new JsonObject
            {
                ["id"] = definition.TryGetProperty("id", out var id) ? id.GetString() : $"series-{definitionIndex + 1}",
                ["name"] = definition.TryGetProperty("name", out var name) ? name.GetString() : InferSeriesName(encode, columnNames, type, definitionIndex),
                ["values"] = new JsonArray(),
            };
            if (definition.TryGetProperty("xAxisIndex", out var authoredXAxisIndex))
                item["xAxisIndex"] = authoredXAxisIndex.GetInt32();
            if (definition.TryGetProperty("yAxisIndex", out var authoredYAxisIndex))
                item["yAxisIndex"] = authoredYAxisIndex.GetInt32();
            var values = (JsonArray)item["values"]!;
            if (type is "scatter" or "bubble")
            {
                if (xColumn is null) throw new InvalidOperationException($"Dataset series {definitionIndex + 1} requires x for {type}.");
                item["xValues"] = new JsonArray(rows.Select(row => NumberNode(row[xColumn.Value])).ToArray());
                foreach (var row in rows) values.Add(NumberNode(row[valueColumn!.Value]));
                if (type == "bubble")
                {
                    var size = RefIndex(encode, "size", columnNames) ?? throw new InvalidOperationException("Bubble encoding requires size.");
                    item["bubbleSizes"] = new JsonArray(rows.Select(row => NumberNode(row[size])).ToArray());
                }
            }
            else if (type == "sankey")
            {
                var source = RefIndex(encode, "source", columnNames) ?? throw new InvalidOperationException("Sankey encoding requires source.");
                var target = RefIndex(encode, "target", columnNames) ?? throw new InvalidOperationException("Sankey encoding requires target.");
                item["sources"] = new JsonArray(rows.Select(row => ValueNode(row[source])).ToArray());
                item["targets"] = new JsonArray(rows.Select(row => ValueNode(row[target])).ToArray());
                foreach (var row in rows) values.Add(NumberNode(row[valueColumn!.Value]));
            }
            else
            {
                foreach (var category in categories)
                {
                    var row = categoryColumn is { } categoryIndex
                        ? rows.FirstOrDefault(candidate => ScalarKey(candidate[categoryIndex]) == ScalarKey(category))
                        : rows.ElementAtOrDefault(Array.IndexOf(categories, category));
                    values.Add(row is null ? null : NumberNode(row[valueColumn!.Value]));
                }
                AddEncodedArray(item, "openValues", rows, RefIndex(encode, "open", columnNames), false, categories, categoryColumn, NumberNode);
                AddEncodedArray(item, "highValues", rows, RefIndex(encode, "high", columnNames), false, categories, categoryColumn, NumberNode);
                AddEncodedArray(item, "lowValues", rows, RefIndex(encode, "low", columnNames), false, categories, categoryColumn, NumberNode);
                AddEncodedArray(item, "parents", rows, RefIndex(encode, "parent", columnNames), false, categories, categoryColumn, ValueNode);
                var isTotal = RefIndex(encode, "isTotal", columnNames);
                if (isTotal is { })
                    item["pointRoles"] = new JsonArray(categories.Select(category =>
                    {
                        var row = categoryColumn is { } categoryIndex
                            ? rows.FirstOrDefault(candidate => ScalarKey(candidate[categoryIndex]) == ScalarKey(category))
                            : rows.ElementAtOrDefault(Array.IndexOf(categories, category));
                        return (JsonNode?)JsonValue.Create(row is not null && row[isTotal.Value].ValueKind == JsonValueKind.True ? "total" : "delta");
                    }).ToArray());
            }
            if ((type is "treemap" or "sunburst") && definition.TryGetProperty("levels", out var levels))
                item["levels"] = levels.GetInt32();
            var effectiveType = (definition.TryGetProperty("chartType", out var explicitChartType) ? explicitChartType.GetString() : type == "bar" ? "bar" : type) ?? type;
            if (chartType == "combo" || (chartType == "candlestick" && definitionIndex > 0)) item["chartType"] = effectiveType;
            if (indexedSecondary) item["axis"] = "secondary";
            else if (definition.TryGetProperty("axis", out var axis)) item["axis"] = axis.GetString();
            foreach (var property in new[] { "fill", "stroke", "color", "marker", "trendlines", "errorBars", "dataLabels" })
                if (definition.TryGetProperty(property, out var value)) item[property] = ToNode(value);
            ApplySeriesDefaults(
                item,
                data.TryGetProperty("seriesDefaults", out var defaults) ? defaults : default,
                effectiveType,
                type);
            outputSeries.Add(item);
        }
        var output = new JsonObject
        {
            ["categories"] = new JsonArray(categories.Select(ToNode).ToArray()),
            ["series"] = outputSeries,
        };
        return JsonDocument.Parse(output.ToJsonString()).RootElement.Clone();
    }

    private static JsonElement CanonicalHeatmapDatasetSeries(
        JsonElement data,
        IReadOnlyList<string> columns,
        IReadOnlyList<JsonElement[]> allRows,
        IReadOnlyList<JsonElement> definitions)
    {
        var definition = definitions.Single(item => item.GetProperty("type").GetString() == "heatmap");
        var encode = definition.GetProperty("encode");
        var xColumn = RefIndex(encode, "x", columns) ?? throw new InvalidOperationException("Heatmap encoding requires x.");
        var yColumn = RefIndex(encode, "y", columns) ?? throw new InvalidOperationException("Heatmap encoding requires y.");
        var valueColumn = RefIndex(encode, "value", columns) ?? throw new InvalidOperationException("Heatmap encoding requires value.");
        var categories = allRows.Select(row => row[xColumn]).GroupBy(ScalarKey, StringComparer.Ordinal).Select(group => group.First()).ToArray();
        var names = allRows.Select(row => row[yColumn]).GroupBy(ScalarKey, StringComparer.Ordinal).Select(group => group.First()).ToArray();
        var series = new JsonArray();
        foreach (var name in names)
        {
            var item = new JsonObject
            {
                ["id"] = $"series-{series.Count + 1}",
                ["name"] = DisplayText(name),
                ["values"] = new JsonArray(),
            };
            var values = (JsonArray)item["values"]!;
            foreach (var category in categories)
            {
                var row = allRows.FirstOrDefault(candidate => ScalarKey(candidate[xColumn]) == ScalarKey(category) && ScalarKey(candidate[yColumn]) == ScalarKey(name));
                values.Add(row is null ? null : NumberNode(row[valueColumn]));
            }
            series.Add(item);
        }
        return JsonDocument.Parse(new JsonObject
        {
            ["categories"] = new JsonArray(categories.Select(ToNode).ToArray()),
            ["series"] = series,
        }.ToJsonString()).RootElement.Clone();
    }

    private static IReadOnlyList<JsonElement[]> ApplyFilters(
        IReadOnlyList<JsonElement[]> rows,
        JsonElement filters,
        IReadOnlyList<string> columns)
    {
        if (filters.ValueKind != JsonValueKind.Array) return rows;
        return rows.Where(row => filters.EnumerateArray().All(filter => MatchesFilter(row, filter, columns))).ToArray();
    }

    private static string InferSeriesName(JsonElement encode, IReadOnlyList<string> columns, string type, int index)
    {
        foreach (var channel in new[] { "y", "value", "close", "flow", "category" })
            if (RefIndex(encode, channel, columns) is { } column) return columns[column];
        return $"{type} {index + 1}";
    }

    private static void ApplySeriesDefaults(
        JsonObject item,
        JsonElement defaults,
        string effectiveType,
        string? sourceType = null)
    {
        if (defaults.ValueKind != JsonValueKind.Object) return;
        JsonElement selected = default;
        var selectedTyped = false;
        foreach (var candidate in new[] { sourceType, effectiveType, effectiveType == "column" ? "bar" : null })
        {
            if (candidate is not null && defaults.TryGetProperty(candidate, out var typed) && typed.ValueKind == JsonValueKind.Object)
            {
                selected = typed;
                selectedTyped = true;
                break;
            }
        }
        if (selected.ValueKind != JsonValueKind.Object)
            selected = defaults;

        foreach (var property in selected.EnumerateObject())
        {
            // A generic fallback may contain the grouped defaults for other
            // series types. Those keys are not series fields themselves.
            if (!selectedTyped && IsSeriesDefaultType(property.Name)) continue;
            var defaultValue = ToNode(property.Value);
            if (!item.TryGetPropertyValue(property.Name, out var existing) || existing is null)
            {
                item[property.Name] = defaultValue;
                continue;
            }
            if (defaultValue is JsonObject defaultObject && existing is JsonObject existingObject)
            {
                item[property.Name] = MergeSeriesDefaultObjects(defaultObject, existingObject);
            }
        }
    }

    private static JsonObject MergeSeriesDefaultObjects(JsonObject defaults, JsonObject overrides)
    {
        var merged = (JsonObject)defaults.DeepClone();
        foreach (var child in overrides)
        {
            if (child.Value is JsonObject overrideObject && merged[child.Key] is JsonObject defaultObject)
            {
                merged[child.Key] = MergeSeriesDefaultObjects(defaultObject, overrideObject);
            }
            else
            {
                merged[child.Key] = child.Value?.DeepClone();
            }
        }
        return merged;
    }

    private static bool IsSeriesDefaultType(string property) => property is
        "bar" or "line" or "area" or "scatter" or "bubble" or "candlestick" or
        "pie" or "radar" or "waterfall" or "heatmap" or "treemap" or "sunburst" or "sankey";

    private static JsonElement[] DatasetRow(JsonElement row, IReadOnlyList<string> columns)
    {
        if (row.ValueKind == JsonValueKind.Array)
        {
            var values = row.EnumerateArray().Select(item => item.Clone()).ToArray();
            if (values.Length != columns.Count)
                throw new InvalidOperationException($"Chart dataset row has {values.Length} values for {columns.Count} columns.");
            return values;
        }
        if (row.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException("Chart dataset rows must be arrays or objects.");
        var columnSet = columns.ToHashSet(StringComparer.Ordinal);
        if (row.EnumerateObject().Any(property => !columnSet.Contains(property.Name)))
            throw new InvalidOperationException("Chart dataset row contains a column that is not declared in cols.");
        return columns.Select(column => row.TryGetProperty(column, out var value) ? value.Clone() : JsonDocument.Parse("null").RootElement.Clone()).ToArray();
    }

    private static int? RefIndex(JsonElement owner, string property, IReadOnlyList<string> columns)
    {
        if (!owner.TryGetProperty(property, out var value)) return null;
        if (value.ValueKind == JsonValueKind.Number) return value.GetInt32();
        var name = value.GetString();
        var index = Array.FindIndex(columns.ToArray(), column => column.Equals(name, StringComparison.Ordinal));
        return index >= 0 ? index : throw new InvalidOperationException($"Unknown chart dataset column {name}.");
    }

    private static void AddEncodedArray(
        JsonObject item,
        string name,
        IReadOnlyList<JsonElement[]> rows,
        int? column,
        bool numericX,
        IReadOnlyList<JsonElement> categories,
        int? categoryColumn,
        Func<JsonElement, JsonNode?> convert)
    {
        if (column is null) return;
        var output = new JsonArray();
        if (numericX) foreach (var row in rows) output.Add(convert(row[column.Value]));
        else foreach (var category in categories)
        {
            var row = categoryColumn is { } index
                ? rows.FirstOrDefault(candidate => ScalarKey(candidate[index]) == ScalarKey(category))
                : rows.ElementAtOrDefault(Array.IndexOf(categories.ToArray(), category));
            output.Add(row is null ? null : convert(row[column.Value]));
        }
        item[name] = output;
    }

    private static bool MatchesFilter(JsonElement[] row, JsonElement filter, IReadOnlyList<string> columns)
    {
        var column = RefIndex(filter, "column", columns) ?? throw new InvalidOperationException("Chart filter requires a column.");
        var actual = row[column];
        var operation = filter.GetProperty("op").GetString()!;
        var expected = filter.GetProperty("value");
        if (operation == "in") return expected.ValueKind == JsonValueKind.Array && expected.EnumerateArray().Any(value => Compare(actual, value) == 0);
        var comparison = Compare(actual, expected);
        return operation switch
        {
            "eq" => comparison == 0,
            "neq" => comparison != 0,
            "gt" => comparison > 0,
            "gte" => comparison >= 0,
            "lt" => comparison < 0,
            "lte" => comparison <= 0,
            _ => throw new InvalidOperationException($"Unsupported chart data filter operator {operation}."),
        };
    }

    private static int Compare(JsonElement left, JsonElement right)
    {
        if (left.ValueKind == JsonValueKind.Number && right.ValueKind == JsonValueKind.Number)
            return left.GetDouble().CompareTo(right.GetDouble());
        return string.Compare(ScalarKey(left), ScalarKey(right), StringComparison.Ordinal);
    }

    private static string DisplayText(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.Number => value.GetDouble().ToString("R", System.Globalization.CultureInfo.InvariantCulture),
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        JsonValueKind.Object when value.TryGetProperty("date", out var date) => date.GetString() ?? "",
        _ => ScalarKey(value),
    };

    private static string ScalarKey(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null => "",
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.Number => value.GetDouble().ToString("R", System.Globalization.CultureInfo.InvariantCulture),
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        JsonValueKind.Object => value.TryGetProperty("date", out var date) ? date.GetString() ?? "" : value.GetRawText(),
        _ => value.GetRawText(),
    };

    private static JsonNode? ValueNode(JsonElement value) => value.ValueKind == JsonValueKind.Undefined ? null : ToNode(value);
    private static JsonNode? ToNode(JsonElement value) => JsonNode.Parse(value.GetRawText());
    private static JsonNode? NumberNode(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Null) return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number) && double.IsFinite(number))
            return JsonValue.Create(number);
        if (value.ValueKind == JsonValueKind.String && double.TryParse(
                value.GetString(),
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out number) && double.IsFinite(number))
            return JsonValue.Create(number);
        throw new InvalidOperationException("Chart dataset numeric channel contains a non-numeric value.");
    }

    private static PpjTableRowModel ParseTableRow(JsonElement row) => new(
        OptionalString(row, "id"),
        row.TryGetProperty("height", out var height) ? height.GetDouble() : null,
        row.GetProperty("cells").EnumerateArray().Select(cell => new PpjTableCellModel(
            OptionalString(cell, "id"),
            ParseText(cell.GetProperty("text")),
            cell.TryGetProperty("rowSpan", out var rowSpan) ? rowSpan.GetInt32() : 1,
            cell.TryGetProperty("columnSpan", out var columnSpan) ? columnSpan.GetInt32() : 1,
            cell.Clone())).ToArray());

    private static PpjConnectorEndpointModel ParseConnectorEndpoint(JsonElement endpoint) => new(
        OptionalString(endpoint, "element"),
        OptionalString(endpoint, "anchor"),
        endpoint.TryGetProperty("x", out var x) ? x.GetDouble() : null,
        endpoint.TryGetProperty("y", out var y) ? y.GetDouble() : null);

    private static PpjSmartArtNodeModel ParseSmartArtNode(JsonElement node) => new(
        node.GetProperty("id").GetString()!,
        ParseText(node.GetProperty("text")),
        OptionalString(node, "styleRef"),
        OptionalString(node, "shapeStyleRef"),
        OptionalString(node, "asset"),
        node.TryGetProperty("image", out var image) ? image.Clone() : null,
        node.TryGetProperty("nativeRef", out var nativeRef) ? ParseNativeRef(nativeRef) : null,
        node.Clone());

    private static PpjSmartArtConnectionModel ParseSmartArtConnection(JsonElement connection) => new(
        connection.GetProperty("id").GetString()!,
        connection.GetProperty("from").GetString()!,
        connection.GetProperty("to").GetString()!,
        connection.GetProperty("role").GetString()!,
        connection.TryGetProperty("order", out var order) ? order.GetUInt32() : 0);

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
                        : null,
                    run.TryGetProperty("field", out var field)
                        ? new PpjTextFieldModel(OptionalString(field, "id"), field.GetProperty("type").GetString()!, field.GetProperty("text").GetString()!)
                        : null,
                    run.TryGetProperty("break", out var lineBreak) && lineBreak.ValueKind == JsonValueKind.True)).ToArray())).ToArray());
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
            layout.ValueKind == JsonValueKind.Object && layout.TryGetProperty("gap", out var gap) ? gap.GetDouble() : 0,
            layout.ValueKind == JsonValueKind.Object && layout.TryGetProperty("columns", out var columns) ? columns.GetInt32() : null,
            layout.ValueKind == JsonValueKind.Object && layout.TryGetProperty("rowGap", out var rowGap) ? rowGap.GetDouble() : null,
            layout.ValueKind == JsonValueKind.Object ? OptionalString(layout, "anchor") : null,
            layout.ValueKind == JsonValueKind.Object && layout.TryGetProperty("weights", out var weights)
                ? weights.EnumerateArray().Select(value => value.GetDouble()).ToArray()
                : null);
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

    private static string? OptionalLiteralString(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

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
