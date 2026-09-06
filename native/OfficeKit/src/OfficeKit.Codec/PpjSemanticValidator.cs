using System.Text;
using System.Text.Json;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static class PpjSemanticValidator
{
    private const double EmuPerPoint = 12_700d;
    private static readonly IReadOnlySet<string> AuthoredVideoMimeTypes = Set("video/mp4");
    private static readonly IReadOnlySet<string> AuthoredAudioMimeTypes = Set("audio/mpeg", "audio/mp4", "audio/wav", "audio/x-wav");
    private static readonly IReadOnlySet<string> FormalTextPrecedenceTargets = Set(
        "text.size", "text.bold", "text.italic", "text.font", "text.fontFamily", "text.fontFamilyEastAsia", "text.fontFamilyComplexScript", "text.language");
    private static readonly IReadOnlySet<string> FormalTextPrecedenceSources = Set(
        "layout", "element", "paragraph", "run");

    private static readonly IReadOnlyDictionary<string, IReadOnlySet<string>> CapabilityFields =
        new Dictionary<string, IReadOnlySet<string>>(StringComparer.Ordinal)
        {
            ["replaceText"] = Set("text", "visibleText"),
            ["setTextParagraphStyle"] = Set("textStyles", "text.paragraphs[].style.alignment", "text.paragraphs[].style.tabStops"),
            ["setFill"] = Set("fill"),
            ["setStroke"] = Set("stroke"),
            ["setLinePath"] = Set("line.path"),
            ["setTextBodyStyle"] = Set("text.style", "textStyle"),
            ["setShapeEffects"] = Set("shape.shadow", "shape.glow", "shape.innerShadow", "shape.reflection", "shape.softEdge"),
            ["setOpacity"] = Set("opacity", "compositing.opacity"),
            ["setFrame"] = Set("frame.x", "frame.y", "frame.width", "frame.height", "frame.rotation", "frame.flipH", "frame.flipV"),
            ["setGeometry"] = Set("geometry.adjustments", "geometry.paths"),
            ["setCanvas"] = Set("canvas.width", "canvas.height"),
            ["setBackground"] = Set("background"),
            ["setTransition"] = Set("transition"),
            ["setNotes"] = Set("notes"),
            ["replaceImage"] = Set("image.asset"),
            ["replaceSvg"] = Set("image.svgAsset"),
            ["setImageCrop"] = Set("image.crop"),
            ["setImageFit"] = Set("image.fit"),
            ["setImageMask"] = Set("image.mask.preset", "image.mask.adjustments", "image.mask.paths"),
            ["setImageEffects"] = Set("image.border", "image.shadow", "image.glow", "image.innerShadow", "image.reflection", "image.softEdge"),
            ["setTableStyle"] = Set("table.style"),
            ["setTableGeometry"] = Set("table.geometry"),
            ["setTableCellStyle"] = Set("table.cell.fill", "table.cell.borders", "table.cell.textStyle"),
            ["setChartTitle"] = Set("chart.title"),
            ["setChartData"] = Set("chart.data"),
            ["setChartTextStyle"] = Set("chart.textStyle"),
            ["setChartFill"] = Set("chart.fill", "chart.legendFill"),
            ["setChartSeriesStyle"] = Set("chart.data.series[].stroke", "chart.data.series[].marker"),
            ["setChartSeriesAnalytics"] = Set("chart.data.series[].trendlines", "chart.data.series[].errorBars"),
            ["setChartFrame"] = Set("chart.frame"),
            ["setChartLabels"] = Set("chart.labels"),
            ["setChartAxis"] = Set("chart.axis"),
            ["setChartPlot"] = Set("chart.plot"),
            ["setAction"] = Set("action"),
            ["setHoverAction"] = Set("hoverAction"),
            ["setCommentStatus"] = Set("status", "resolved"),
            ["setSmartArtText"] = Set("smartArt.text"),
            ["setSmartArtImage"] = Set("smartArt.nodes[].asset"),
            ["setSmartArtImagePaint"] = Set("smartArt.nodes[].image"),
            ["setOlePayload"] = Set("ole.payload"),
            ["setName"] = Set("name"),
            ["setPages"] = Set("pages"),
            ["setHidden"] = Set("hidden"),
            ["setLocked"] = Set("locked"),
            ["appendElement"] = Set("elements"),
            ["delete"] = Set("element"),
            ["duplicate"] = Set("element", "pageClone"),
            ["reorder"] = Set("zOrder", "pageOrder"),
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
        var colorIds = program.Design.ColorIds.ToHashSet(StringComparer.Ordinal);
        var grammarTokenKinds = new Dictionary<string, string>(StringComparer.Ordinal);
        if (program.Root.GetProperty("design").TryGetProperty("grammar", out var grammar) &&
            grammar.ValueKind == JsonValueKind.Object &&
            grammar.TryGetProperty("tokens", out var grammarTokens) &&
            grammarTokens.ValueKind == JsonValueKind.Object)
        {
            foreach (var token in grammarTokens.EnumerateObject())
            {
                if (token.Value.ValueKind == JsonValueKind.Object &&
                    token.Value.TryGetProperty("kind", out var kind) &&
                    kind.GetString() is { } kindName)
                {
                    grammarTokenKinds[token.Name] = kindName;
                    if (kindName == "color") colorIds.Add(token.Name);
                }
            }
        }
        ValidateResourceReferences(program.Root, new StringBuilder("$"), assetIds, colorIds, program.Design.FontIds, grammarTokenKinds, diagnostics);
        ValidateDesignGrammar(program.Root.GetProperty("design"), "$.design.grammar", diagnostics);
        ValidateFormalTextStyleBoundary(program, diagnostics);
        ValidateTextEffects(program.Root, new StringBuilder("$"), diagnostics);
        ValidateFormulaRuns(program.Root, new StringBuilder("$"), diagnostics);
        ValidateMasterLayoutState(program, masters, layouts, diagnostics);
        ValidateComponentDefinitions(program, components, assetIds, diagnostics);
        ValidateNativeRef(program.Design.CanvasNativeRef, program.Source, "$.design.canvas.nativeRef", diagnostics);

        var globalElementIds = new HashSet<string>(StringComparer.Ordinal);
        var globalAnimationIds = new HashSet<string>(StringComparer.Ordinal);
        var pageElements = new Dictionary<string, IReadOnlyDictionary<string, PpjElementModel>>(StringComparer.Ordinal);
        var clonedSourcePageIds = new HashSet<string>(StringComparer.Ordinal);

        for (var pageIndex = 0; pageIndex < program.Pages.Count; pageIndex++)
        {
            var page = program.Pages[pageIndex];
            var pagePath = $"$.pages[{pageIndex}]";
            ValidateNativeRef(page.NativeRef, program.Source, $"{pagePath}.nativeRef", diagnostics);
            ValidateSourceClone(program, pages, page, pageIndex, pagePath, clonedSourcePageIds, diagnostics);

            var localElements = IndexElements(
                page.Elements,
                $"{pagePath}.elements",
                globalElementIds,
                diagnostics);
            pageElements[page.Id] = localElements;
            // A source-bound page may still carry the previous direct-order
            // permutation while the request omits a source element for a
            // capability-issued deletion.  The source compiler filters that
            // proven stale ID against the exact baseline; source-free PPJ
            // remains strict and must provide a complete permutation.
            ValidateReadingOrder(page, pagePath, program.Source is not null, diagnostics);

            foreach (var (element, path) in WalkElements(page.Elements, $"{pagePath}.elements"))
                ValidateElement(element, path, program, components, assetIds, localElements, inComponent: false, diagnostics);

            ValidateTimingGraph(page.Raw, pagePath, diagnostics);
            ValidateAnimations(page, pagePath, localElements, globalAnimationIds, diagnostics);
        }

        ValidateTransitions(program.Pages, pages, pageElements, diagnostics);
        ValidatePresentationReferences(program, pageIds, pageElements, diagnostics);
        ValidateComponentGraph(program.Components, diagnostics);
        ValidateExpansionBudget(program, components, diagnostics);
    }

    private static void ValidateSourceClone(
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjPageModel> pages,
        PpjPageModel page,
        int pageIndex,
        string path,
        ISet<string> clonedSourcePageIds,
        List<PpjDiagnostic> diagnostics)
    {
        var clone = page.SourceClone;
        if (clone is null) return;
        if (program.Source is null)
            diagnostics.Add(new("ppj.sourceClone.source", "sourceClone is only valid in a source-bound PPJ.", path + ".sourceClone"));
        if (page.NativeRef is not null)
            diagnostics.Add(new("ppj.sourceClone.nativeRef", "A pending sourceClone page cannot carry its own nativeRef before build and reimport.", path + ".nativeRef"));
        if (page.Elements.Count != 0)
            diagnostics.Add(new("ppj.sourceClone.elements", "A pending sourceClone page must keep elements empty until build and reimport.", path + ".elements"));
        if (page.Name is not null || page.LayoutId is not null || page.Notes is not null || page.Hidden is not null ||
            page.Transition is not null || page.Animations.Count != 0 || page.Raw.TryGetProperty("background", out _))
            diagnostics.Add(new(
                "ppj.sourceClone.immutable",
                "A pending sourceClone cannot declare name, layout, background, notes, visibility, transition, or animation state before build and reimport.",
                path));
        if (!pages.TryGetValue(clone.PageId, out var sourcePage))
        {
            diagnostics.Add(new("ppj.sourceClone.page", $"Source page {clone.PageId} does not exist.", path + ".sourceClone.page"));
            return;
        }
        if (sourcePage.SourceClone is not null)
            diagnostics.Add(new("ppj.sourceClone.chain", "A pending sourceClone cannot clone another pending clone.", path + ".sourceClone.page"));
        if (pageIndex == 0 || !program.Pages[pageIndex - 1].Id.Equals(clone.PageId, StringComparison.Ordinal))
            diagnostics.Add(new("ppj.sourceClone.adjacent", "A pending sourceClone must immediately follow its retained source page.", path + ".sourceClone.page"));
        if (!clonedSourcePageIds.Add(clone.PageId))
            diagnostics.Add(new("ppj.sourceClone.budget", $"Source page {clone.PageId} can be cloned only once before build and reimport.", path + ".sourceClone.page"));
        var capability = sourcePage.NativeRef?.Capabilities.FirstOrDefault(item => item.Id.Equals(clone.CapabilityId, StringComparison.Ordinal));
        if (capability is null || !capability.Operation.Equals("duplicate", StringComparison.Ordinal) ||
            !capability.Fields.Contains("pageClone", StringComparer.Ordinal) ||
            !capability.ExpectedHash.Equals(sourcePage.NativeRef!.ObjectHash, StringComparison.OrdinalIgnoreCase))
            diagnostics.Add(new(
                "ppj.sourceClone.capability",
                "The referenced source page did not issue this duplicate/pageClone capability.",
                path + ".sourceClone.capability"));

        if (clone.RetainElementId is not { } retainedElementId) return;
        if (!sourcePage.Elements.Any(element => element.Id.Equals(retainedElementId, StringComparison.Ordinal)))
        {
            diagnostics.Add(new(
                "ppj.sourceClone.retainElement",
                $"Source page {clone.PageId} has no direct element {retainedElementId}.",
                path + ".sourceClone.retainElement"));
            return;
        }

        foreach (var sibling in sourcePage.Elements.Where(element => !element.Id.Equals(retainedElementId, StringComparison.Ordinal)))
        {
            var deletion = sibling.NativeRef?.Capabilities.FirstOrDefault(item =>
                item.Operation.Equals("delete", StringComparison.Ordinal) &&
                item.Fields.Contains("element", StringComparer.Ordinal));
            if (deletion is not null &&
                deletion.ExpectedHash.Equals(sibling.NativeRef!.ObjectHash, StringComparison.OrdinalIgnoreCase)) continue;
            diagnostics.Add(new(
                "ppj.sourceClone.siblingDelete",
                $"Source element {sibling.Id} did not issue the delete/element capability required by component reuse.",
                path + ".sourceClone.retainElement"));
        }
    }

    private static void ValidateReadingOrder(
        PpjPageModel page,
        string path,
        bool sourceBound,
        List<PpjDiagnostic> diagnostics)
    {
        ValidateReadingOrder(page.Elements, page.ReadingOrder, path, sourceBound, diagnostics);
    }

    private static void ValidateReadingOrder(
        IReadOnlyList<PpjElementModel> elements,
        IReadOnlyList<string> readingOrder,
        string path,
        bool sourceBound,
        List<PpjDiagnostic> diagnostics)
    {
        if (readingOrder.Count == 0) return;
        var directIds = elements.Select(element => element.Id).ToArray();
        var expected = directIds.ToHashSet(StringComparer.Ordinal);
        var actual = readingOrder.ToHashSet(StringComparer.Ordinal);
        // Source-bound requests can be based on a projection made before a
        // capability-issued deletion or append.  The compiler re-proves
        // those stale/missing IDs against the exact source page; the syntax
        // validator only rejects duplicate entries here.  Source-free PPJ
        // has no such baseline and therefore stays a strict permutation.
        var invalid = actual.Count != readingOrder.Count ||
            (sourceBound ? false : !expected.SetEquals(actual));
        if (invalid)
        {
            diagnostics.Add(new(
                "ppj.accessibility.readingOrder",
                "readingOrder must be a complete permutation of the owning container's direct element IDs; it is not inferred from z-order or nested component expansion.",
                path + ".readingOrder"));
        }
    }

    private static void ValidateMasterLayoutState(
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjMasterModel> masters,
        IReadOnlyDictionary<string, PpjLayoutModel> layouts,
        List<PpjDiagnostic> diagnostics)
    {
        for (var masterIndex = 0; masterIndex < program.Design.Masters.Count; masterIndex++)
        {
            var master = program.Design.Masters[masterIndex];
            var path = $"$.design.masters[{masterIndex}]";
            ValidateTextStyleLevels(master.TitleTextLevels, path + ".textStyles.title", diagnostics);
            ValidateTextStyleLevels(master.BodyTextLevels, path + ".textStyles.body", diagnostics);
            ValidateTextStyleLevels(master.OtherTextLevels, path + ".textStyles.other", diagnostics);
            ValidateLayoutPlaceholders(master.Placeholders, path + ".placeholders", program.Source, diagnostics);
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
            ValidateLayoutPlaceholders(layout.Placeholders, path + ".placeholders", program.Source, diagnostics);
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

    private static void ValidateDesignGrammar(
        JsonElement design,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (!design.TryGetProperty("grammar", out var grammar) || grammar.ValueKind != JsonValueKind.Object) return;

        if (grammar.TryGetProperty("tokens", out var tokens) && tokens.ValueKind == JsonValueKind.Object)
        {
            foreach (var token in tokens.EnumerateObject())
            {
                var tokenPath = $"{path}.tokens.{token.Name}";
                var definition = token.Value;
                if (definition.ValueKind != JsonValueKind.Object ||
                    !definition.TryGetProperty("kind", out var kindValue) ||
                    !definition.TryGetProperty("value", out var value)) continue;
                var kind = kindValue.GetString();
                var valid = kind switch
                {
                    "boolean" => value.ValueKind is JsonValueKind.True or JsonValueKind.False,
                    "color" => value.ValueKind == JsonValueKind.String && IsGrammarColor(value.GetString()!),
                    "font" or "string" => value.ValueKind == JsonValueKind.String,
                    "size" or "spacing" or "radius" or "opacity" => value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number) && double.IsFinite(number),
                    _ => false,
                };
                if (!valid)
                {
                    diagnostics.Add(new(
                        "ppj.grammar.tokenValue",
                        $"Grammar token {token.Name} does not match its declared kind {kind ?? "(missing)"}.",
                        $"{tokenPath}.value"));
                    continue;
                }
                if (kind == "opacity" && value.GetDouble() is < 0 or > 1)
                    diagnostics.Add(new("ppj.grammar.opacity", $"Grammar opacity token {token.Name} must be between 0 and 1.", $"{tokenPath}.value"));
                if (kind is "size" or "spacing" or "radius" && value.GetDouble() < 0)
                    diagnostics.Add(new("ppj.grammar.nonNegative", $"Grammar {kind} token {token.Name} cannot be negative.", $"{tokenPath}.value"));
            }
        }

        if (grammar.TryGetProperty("stylePrecedence", out var precedence) && precedence.ValueKind == JsonValueKind.Array)
        {
            var targets = new HashSet<string>(StringComparer.Ordinal);
            var index = 0;
            foreach (var rule in precedence.EnumerateArray())
            {
                var rulePath = $"{path}.stylePrecedence[{index++}]";
                if (!rule.TryGetProperty("target", out var target) || !targets.Add(target.GetString()!))
                    diagnostics.Add(new("ppj.grammar.precedenceTarget", "Each style-precedence target must be declared exactly once.", $"{rulePath}.target"));
                if (rule.TryGetProperty("sources", out var sources))
                {
                    var values = sources.EnumerateArray().Select(item => item.GetString()!).ToArray();
                    if (values.Length == 0 || values.Distinct(StringComparer.Ordinal).Count() != values.Length)
                        diagnostics.Add(new("ppj.grammar.precedenceSources", "Style-precedence sources must be non-empty and unique.", $"{rulePath}.sources"));
                    var targetName = rule.TryGetProperty("target", out var targetValue)
                        ? targetValue.GetString()
                        : null;
                    foreach (var source in values)
                    {
                        if (!FormalTextPrecedenceSources.Contains(source)) continue;
                        if (targetName is null || !FormalTextPrecedenceTargets.Contains(targetName))
                            diagnostics.Add(new(
                                "ppj.grammar.precedenceSourceTarget",
                                $"Style-precedence source {source} is only supported for the bounded text scalar targets.",
                                $"{rulePath}.sources"));
                    }
                }
            }
        }

        if (grammar.TryGetProperty("predicates", out var predicates) && predicates.ValueKind == JsonValueKind.Array)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var index = 0;
            foreach (var predicate in predicates.EnumerateArray())
            {
                var predicatePath = $"{path}.predicates[{index++}]";
                if (predicate.TryGetProperty("id", out var id) && !ids.Add(id.GetString()!))
                    diagnostics.Add(new("ppj.grammar.predicateId", "Grammar predicate IDs must be unique.", $"{predicatePath}.id"));
                if (!predicate.TryGetProperty("value", out var value)) continue;
                var op = predicate.TryGetProperty("op", out var operatorValue) ? operatorValue.GetString() : null;
                var values = value.ValueKind == JsonValueKind.Array ? value.EnumerateArray().ToArray() : [value];
                if (op == "in" && value.ValueKind != JsonValueKind.Array)
                    diagnostics.Add(new("ppj.grammar.predicateIn", "The in predicate operator requires an array of values.", $"{predicatePath}.value"));
                if (op is "gt" or "gte" or "lt" or "lte" && values.Any(item => item.ValueKind != JsonValueKind.Number || !item.TryGetDouble(out var number) || !double.IsFinite(number)))
                    diagnostics.Add(new("ppj.grammar.predicateNumber", $"Predicate operator {op} requires finite numeric values.", $"{predicatePath}.value"));
            }
        }
    }

    private static void ValidateFormalTextStyleBoundary(
        PpjProgramModel program,
        List<PpjDiagnostic> diagnostics)
    {
        if (program.Source is null) return;

        if (program.Root.GetProperty("design").TryGetProperty("grammar", out var grammar) &&
            grammar.TryGetProperty("stylePrecedence", out var precedence) &&
            precedence.ValueKind == JsonValueKind.Array)
        {
            for (var index = 0; index < precedence.GetArrayLength(); index++)
            {
                var rule = precedence[index];
                var target = rule.TryGetProperty("target", out var targetValue) ? targetValue.GetString() : null;
                if (target is null || !target.StartsWith("text.", StringComparison.Ordinal) ||
                    !rule.TryGetProperty("sources", out var sources) || sources.ValueKind != JsonValueKind.Array)
                    continue;
                if (sources.EnumerateArray().Any(source => source.GetString() is "layout" or "master" or "element" or "paragraph" or "run"))
                    diagnostics.Add(new(
                        "ppj.sourceBound.formalTextPrecedence",
                        "Source-bound PPJ cannot declare formal layout, master, element, paragraph, or run text precedence owners.",
                        $"$.design.grammar.stylePrecedence[{index}].sources"));
            }
        }

        var design = program.Root.GetProperty("design");
        RejectSourceBoundOwnerField(design, "masters", "style", "$.design.masters", diagnostics);
        RejectSourceBoundOwnerField(design, "layouts", "style", "$.design.layouts", diagnostics);
        if (design.TryGetProperty("theme", out var theme) &&
            theme.ValueKind == JsonValueKind.Object &&
            theme.TryGetProperty("textStyle", out _))
        {
            diagnostics.Add(new(
                "ppj.sourceBound.formalTextOwner",
                "Source-bound PPJ cannot declare design.theme.textStyle as a new formal text owner.",
                "$.design.theme.textStyle"));
        }
        if (design.TryGetProperty("theme", out theme) &&
            theme.ValueKind == JsonValueKind.Object &&
            theme.TryGetProperty("fontScheme", out _))
        {
            diagnostics.Add(new(
                "ppj.sourceBound.themeFontScheme",
                "Source-bound PPJ cannot declare a new authored design.theme.fontScheme over the preserved native theme graph.",
                "$.design.theme.fontScheme"));
        }
        if (design.TryGetProperty("theme", out theme) &&
            theme.ValueKind == JsonValueKind.Object &&
            theme.TryGetProperty("accentColors", out _))
        {
            diagnostics.Add(new(
                "ppj.sourceBound.themeAccentColors",
                "Source-bound PPJ cannot declare new authored design.theme.accentColors over the preserved native theme graph.",
                "$.design.theme.accentColors"));
        }
        if (design.TryGetProperty("theme", out theme) &&
            theme.ValueKind == JsonValueKind.Object &&
            theme.TryGetProperty("colorRoles", out _))
        {
            diagnostics.Add(new(
                "ppj.sourceBound.themeColorRoles",
                "Source-bound PPJ cannot declare new authored design.theme.colorRoles over the preserved native theme graph.",
                "$.design.theme.colorRoles"));
        }
        foreach (var page in program.Pages.Select((value, index) => (value, index)))
        {
            foreach (var (element, path) in WalkElements(page.value.Elements, $"$.pages[{page.index}].elements"))
            {
                if (element.Type is not ("text" or "placeholder") || !element.Raw.TryGetProperty("textStyle", out _)) continue;
                diagnostics.Add(new(
                    "ppj.sourceBound.formalTextOwner",
                    "Source-bound PPJ cannot declare a new direct textStyle owner on text or placeholder elements.",
                    path + ".textStyle"));
            }
        }
        foreach (var component in program.Components.Select((value, index) => (value, index)))
        {
            foreach (var (element, path) in WalkElements(component.value.Elements, $"$.components[{component.index}].elements"))
            {
                if (element.Type is not ("text" or "placeholder") || !element.Raw.TryGetProperty("textStyle", out _)) continue;
                diagnostics.Add(new(
                    "ppj.sourceBound.formalTextOwner",
                    "Source-bound PPJ cannot declare a new direct textStyle owner on text or placeholder elements.",
                    path + ".textStyle"));
            }
        }
    }

    private static void RejectSourceBoundOwnerField(
        JsonElement design,
        string collection,
        string property,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (!design.TryGetProperty(collection, out var definitions) || definitions.ValueKind != JsonValueKind.Array) return;
        for (var index = 0; index < definitions.GetArrayLength(); index++)
        {
            if (!definitions[index].TryGetProperty(property, out _)) continue;
            diagnostics.Add(new(
                "ppj.sourceBound.formalTextOwner",
                $"Source-bound PPJ cannot declare design.{collection}[].{property} as a new formal text owner.",
                $"{path}[{index}].{property}"));
        }
    }

    private static bool IsGrammarColor(string value)
    {
        var normalized = value.TrimStart('#');
        return (normalized.Length is 6 or 8) && normalized.All(Uri.IsHexDigit);
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
        PpjSourceModel? source,
        List<PpjDiagnostic> diagnostics)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var identities = new HashSet<(string Type, uint Index)>();
        for (var index = 0; index < placeholders.Count; index++)
        {
            var placeholder = placeholders[index];
            ValidateNativeRef(placeholder.NativeRef, source, $"{path}[{index}].nativeRef", diagnostics);
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

    private static void ValidateTextEffects(JsonElement value, StringBuilder path, List<PpjDiagnostic> diagnostics)
    {
        if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                var length = AppendIndex(path, index++);
                ValidateTextEffects(item, path, diagnostics);
                path.Length = length;
            }
            return;
        }
        if (value.ValueKind != JsonValueKind.Object) return;

        var hasColor = value.TryGetProperty("color", out _);
        if (value.TryGetProperty("gradient", out var gradient))
        {
            if (hasColor)
                diagnostics.Add(new("ppj.text.paintConflict", "Text style cannot declare both color and gradient.", path.ToString()));
            if (gradient.TryGetProperty("kind", out var kind) && kind.GetString() == "radial" && gradient.TryGetProperty("angle", out _))
                diagnostics.Add(new("ppj.text.radialAngle", "Radial text gradients cannot declare a linear angle.", PathWithProperties(path, "gradient", "angle")));
            if (gradient.TryGetProperty("stops", out var stops))
            {
                var previous = double.NegativeInfinity;
                var stopIndex = 0;
                foreach (var stop in stops.EnumerateArray())
                {
                    var offset = stop.GetProperty("offset").GetDouble();
                    if (offset < previous)
                        diagnostics.Add(new("ppj.text.gradientOrder", "Text gradient stop offsets must be ordered.", PathWithGradientStopOffset(path, stopIndex)));
                    previous = offset;
                    stopIndex++;
                }
            }
        }

        foreach (var property in value.EnumerateObject())
        {
            var length = AppendProperty(path, property.Name);
            ValidateTextEffects(property.Value, path, diagnostics);
            path.Length = length;
        }
    }

    private static void ValidateFormulaRuns(JsonElement value, StringBuilder path, List<PpjDiagnostic> diagnostics)
    {
        if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                var length = AppendIndex(path, index++);
                ValidateFormulaRuns(item, path, diagnostics);
                path.Length = length;
            }
            return;
        }
        if (value.ValueKind != JsonValueKind.Object) return;

        if (value.TryGetProperty("formula", out var formula))
        {
            var sourcePath = PathWithProperties(path, "formula", "source");
            try
            {
                _ = PpjLatexCompiler.Compile(formula.GetProperty("source").GetString()!, sourcePath);
            }
            catch (CodecException error)
            {
                diagnostics.Add(new(error.Code, error.Message, error.SourcePath ?? sourcePath));
            }

            if (value.TryGetProperty("hyperlink", out _))
                diagnostics.Add(new("ppj.formula.hyperlink", "Formula runs cannot carry hyperlinks.", PathWithProperty(path, "hyperlink")));
            if (value.TryGetProperty("style", out var style))
            {
                foreach (var property in style.EnumerateObject())
                {
                    if (property.Name is "size" or "color") continue;
                    diagnostics.Add(new(
                        "ppj.formula.style",
                        $"Formula runs support only direct size and color; {property.Name} is not supported.",
                        PathWithProperties(path, "style", property.Name)));
                }
            }
        }

        foreach (var property in value.EnumerateObject())
        {
            var length = AppendProperty(path, property.Name);
            ValidateFormulaRuns(property.Value, path, diagnostics);
            path.Length = length;
        }
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

            for (var slotIndex = 0; slotIndex < component.Slots.Count; slotIndex++)
            {
                var slot = component.Slots[slotIndex];
                if (slot.ImagePolicy is not null && !slot.Accepts.Contains("image", StringComparer.Ordinal))
                    diagnostics.Add(new(
                        "ppj.component.imagePolicyType",
                        $"Image policy on slot {slot.Name} requires image in its accepts list.",
                        $"{path}.slots[{slotIndex}].imagePolicy"));
                if (slot.ImagePolicy is { MinimumWidthPx: { } minWidth, MinimumHeightPx: { } minHeight } &&
                    (minWidth <= 0 || minHeight <= 0))
                    diagnostics.Add(new(
                        "ppj.component.imagePolicyBounds",
                        $"Image policy on slot {slot.Name} must use positive minimum pixel dimensions.",
                        $"{path}.slots[{slotIndex}].imagePolicy"));
            }

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
        ValidateCompositing(element, path, diagnostics);
        if ((element.Raw.TryGetProperty("action", out _) || element.Raw.TryGetProperty("hoverAction", out _)) &&
            element is not (PpjTextElementModel or PpjShapeElementModel or PpjIconElementModel or PpjPlaceholderElementModel))
        {
            diagnostics.Add(new(
                "ppj.action.type",
                "Element actions are only supported for shape-producing text, shape, line, icon, and placeholder elements in the bounded profile.",
                path + ".action"));
        }

        switch (element)
        {
            case PpjTextElementModel text:
                ValidateStyleRef(text.StyleRef, program.Design.TextStyleIds, $"{path}.styleRef", diagnostics);
                break;
            case PpjShapeElementModel shape:
                ValidateStyleRef(shape.StyleRef, program.Design.ShapeStyleIds, $"{path}.styleRef", diagnostics);
                if (shape.Type == "line")
                    ValidateLineElement(shape.Raw, shape.Frame, path, diagnostics);
                else
                {
                    ValidatePresetAdjustments(shape.GeometryKind, shape.GeometryPreset, shape.GeometryAdjustments, path + ".geometry", diagnostics);
                }
                if (shape.GeometryKind == "custom")
                    ValidateCustomGeometry(shape.Raw.GetProperty("geometry"), path + ".geometry", diagnostics);
                break;
            case PpjIconElementModel icon:
                ValidateStyleRef(icon.StyleRef, program.Design.ShapeStyleIds, $"{path}.styleRef", diagnostics);
                if (!PpjIconCatalog.Contains(icon.IconName))
                    diagnostics.Add(new(
                        "ppj.icon.unknown",
                        $"PPJ iconName {icon.IconName} is not present in the pinned Font Awesome Free catalog.",
                        $"{path}.iconName"));
                break;
            case PpjImageElementModel image:
                ValidateAssetRef(image.AssetId, assetIds, $"{path}.asset", diagnostics);
                ValidateStyleRef(image.StyleRef, program.Design.ImageStyleIds, $"{path}.styleRef", diagnostics);
                ValidateImageFocus(image.Raw, ResolveImageFit(program.Root, image.Raw, image.Fit, path, diagnostics), path, diagnostics);
                if (image.SvgAssetId is not null)
                {
                    ValidateAssetRef(image.SvgAssetId, assetIds, $"{path}.svgAsset", diagnostics);
                    var svgAsset = program.Assets.FirstOrDefault(asset => asset.Id.Equals(image.SvgAssetId, StringComparison.Ordinal));
                    if (svgAsset is not null && !svgAsset.MimeType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                        diagnostics.Add(new("ppj.image.svgMime", "PPJ image.svgAsset requires MIME image/svg+xml.", $"{path}.svgAsset"));
                    if (program.Source is null)
                        diagnostics.Add(new("ppj.image.svgAssetSource", "PPJ image.svgAsset is reserved for an imported native raster/SVG fallback pair; authored SVG images use asset.", $"{path}.svgAsset"));
                }
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
            case PpjGroupElementModel group:
                // Unlike a page, a group cannot append a topmost authored
                // overlay to a source-owned child tree.  Its explicit order
                // therefore has to name every direct child even in a
                // source-bound request; the compiler will still re-prove
                // each child's native reorder capability.
                ValidateReadingOrder(group.Elements, group.ReadingOrder, path, sourceBound: false, diagnostics);
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
                    program.Assets,
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
                ValidateComponentInstance(instance, path, program, components, diagnostics);
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

    private static void ValidateCompositing(PpjElementModel element, string path, List<PpjDiagnostic> diagnostics)
    {
        if (!element.Raw.TryGetProperty("compositing", out var compositing)) return;
        if (compositing.TryGetProperty("blendMode", out var blendMode) && blendMode.GetString() is not (null or "normal"))
            diagnostics.Add(new(
                "ppj.compositing.blendUnsupported",
                "Only normal compositing is currently representable by the native PowerPoint profile; non-normal blend modes remain explicit unsupported semantics.",
                path + ".compositing.blendMode"));
        if (compositing.TryGetProperty("isolation", out var isolation) && isolation.GetBoolean())
            diagnostics.Add(new(
                "ppj.compositing.isolationUnsupported",
                "Group/layer isolation is not represented by DrawingML in the bounded profile.",
                path + ".compositing.isolation"));
        if (compositing.TryGetProperty("clipStack", out var clipStack) && clipStack.GetArrayLength() > 0 &&
            !TryValidateSingleImageClip(element, compositing, out var clipReason))
            diagnostics.Add(new(
                "ppj.compositing.clipUnsupported",
                $"The compositing clip stack is outside the bounded single-image mask profile: {clipReason}",
                path + ".compositing.clipStack"));
        if (!compositing.TryGetProperty("opacity", out _)) return;
        if (element is not (PpjTextElementModel or PpjShapeElementModel or PpjIconElementModel or PpjImageElementModel or PpjPlaceholderElementModel or PpjConnectorElementModel))
            diagnostics.Add(new(
                "ppj.compositing.opacityUnsupported",
                $"Element type {element.Type} has no bounded native element-opacity owner.",
                path + ".compositing.opacity"));
        if (element.Raw.TryGetProperty("opacity", out _))
            diagnostics.Add(new(
                "ppj.compositing.opacityConflict",
                "Use either the element-specific opacity field or compositing.opacity, not both.",
                path + ".compositing.opacity"));
    }

    private static bool TryValidateSingleImageClip(
        PpjElementModel element,
        JsonElement compositing,
        out string reason)
    {
        if (element is not PpjImageElementModel)
        {
            reason = "only image elements have a bounded native clip owner";
            return false;
        }
        if (element.Raw.TryGetProperty("mask", out _))
        {
            reason = "an image mask cannot be combined with a compositing clip";
            return false;
        }
        var clipStack = compositing.GetProperty("clipStack");
        if (clipStack.GetArrayLength() != 1)
        {
            reason = "exactly one clip entry is supported";
            return false;
        }
        var clip = clipStack[0];
        if (clip.TryGetProperty("inverse", out var inverse) && inverse.GetBoolean())
        {
            reason = "inverse clips have no bounded native owner";
            return false;
        }
        var geometry = clip.GetProperty("geometry");
        var geometryKind = geometry.GetProperty("kind").GetString();
        if (geometryKind == "custom")
        {
            var customMask = new PresentationShape
            {
                Geometry = "custom",
                WidthEmu = checked((long)Math.Round(element.Frame.Width * EmuPerPoint)),
                HeightEmu = checked((long)Math.Round(element.Frame.Height * EmuPerPoint)),
            };
            try
            {
                PpjAuthoredPresentationCompiler.ApplyCustomGeometry(customMask, geometry, element.Id + " compositing clip");
                PptxCustomGeometryCodec.Validate(customMask, element.Id + " compositing clip");
            }
            catch (CodecException exception)
            {
                reason = exception.Message;
                return false;
            }
            reason = string.Empty;
            return true;
        }
        if (geometryKind != "preset")
        {
            reason = "only bounded preset or custom clip geometry is supported";
            return false;
        }
        var preset = geometry.GetProperty("preset").GetString();
        if (preset is null || !PptxPresetGeometryAdjustmentCodec.HasProfile(preset))
        {
            reason = $"preset geometry {preset ?? "(missing)"} has no native mask profile";
            return false;
        }
        var adjustments = geometry.TryGetProperty("adjustments", out var rawAdjustments)
            ? rawAdjustments.EnumerateArray().Select(item => item.GetInt32()).ToArray()
            : [];
        try
        {
            PptxPresetGeometryAdjustmentCodec.Validate(preset, adjustments, element.Id + " compositing clip");
        }
        catch (CodecException exception)
        {
            reason = exception.Message;
            return false;
        }
        reason = string.Empty;
        return true;
    }

    private static void ValidateLineElement(JsonElement raw, PpjFrameModel frame, string path, List<PpjDiagnostic> diagnostics)
    {
        if (!raw.TryGetProperty("path", out var linePath))
        {
            if (!raw.TryGetProperty("points", out _))
            {
                diagnostics.Add(new("ppj.line.path", "A line element requires a structured path or Kimi points.", path + ".path"));
                return;
            }
            try
            {
                _ = PpjLinePathCodec.KimiPath(raw, frame.Width, frame.Height, path);
            }
            catch (CodecException exception)
            {
                diagnostics.Add(new(exception.Code, exception.Message, path + ".points"));
            }
            return;
        }
        if (!raw.TryGetProperty("stroke", out _))
            diagnostics.Add(new("ppj.line.stroke", "A line element requires an explicit stroke.", path + ".stroke"));
        var commands = linePath.TryGetProperty("commands", out var commandArray) && commandArray.ValueKind == JsonValueKind.Array
            ? commandArray.EnumerateArray().ToArray()
            : [];
        if (commands.Length < 2 || !commands[0].TryGetProperty("op", out var firstOp) || firstOp.GetString() != "moveTo")
            diagnostics.Add(new("ppj.line.pathStart", "A line path must start with moveTo and contain at least one drawing command.", path + ".path.commands"));
        for (var index = 0; index < commands.Length; index++)
        {
            if (commands[index].TryGetProperty("op", out var operation) && operation.GetString() == "close")
                diagnostics.Add(new("ppj.line.pathClosed", "A line path cannot close a subpath.", $"{path}.path.commands[{index}]"));
        }
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
        var numericCombo = IsNumericCombo(chart);
        if (chart.Raw.TryGetProperty("displayBlanksAs", out var displayBlanksAs))
        {
            if (chart.ChartType is not ("bar" or "column" or "line" or "area" or "pie" or "doughnut" or "scatter" or "bubble" or "radar" or "combo" or "waterfall"))
                diagnostics.Add(new(
                    "ppj.chart.displayBlanksAsType",
                    "displayBlanksAs applies only to bounded native ChartPart chart families.",
                    path + ".displayBlanksAs"));
            else if (displayBlanksAs.ValueKind == JsonValueKind.String && displayBlanksAs.GetString() is not ("zero" or "gap" or "span"))
                diagnostics.Add(new(
                    "ppj.chart.displayBlanksAsValue",
                    "displayBlanksAs must be zero, gap, or span.",
                    path + ".displayBlanksAs"));
        }
        var seriesIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < chart.Data.Series.Count; index++)
        {
            var series = chart.Data.Series[index];
            var seriesPath = $"{path}.data.series[{index}]";
            var seriesType = chart.ChartType == "combo" ? series.ChartType : chart.ChartType;
            if (!seriesIds.Add(series.Id))
                diagnostics.Add(new("ppj.id.duplicate", $"Chart series ID {series.Id} is duplicated.", $"{seriesPath}.id"));
            if (!numericCombo && seriesType is not ("scatter" or "bubble" or "sankey") && series.Values.Count != chart.Data.Categories.Count)
                diagnostics.Add(new("ppj.chart.lengthMismatch", $"Series {series.Id} has {series.Values.Count} values for {chart.Data.Categories.Count} categories.", $"{seriesPath}.values"));
            if (chart.ChartType != "waterfall" && series.PointRoles.Count != 0)
                diagnostics.Add(new("ppj.chart.pointRoleType", "pointRoles applies only to waterfall charts.", $"{seriesPath}.pointRoles"));
            if (series.Levels is not null && chart.ChartType is not ("treemap" or "sunburst"))
                diagnostics.Add(new("ppj.chart.levelsType", "levels applies only to treemap and sunburst charts.", $"{seriesPath}.levels"));
            if (chart.ChartType == "combo" && string.IsNullOrEmpty(series.ChartType))
                diagnostics.Add(new("ppj.chart.comboSeriesType", "Every combo-chart series requires chartType.", $"{seriesPath}.chartType"));
            var xAxisIndex = series.XAxisIndex;
            var yAxisIndex = series.YAxisIndex;
            if (xAxisIndex is not null || yAxisIndex is not null)
            {
                if (chart.ChartType != "combo")
                    diagnostics.Add(new("ppj.chart.axisIndexType", "xAxisIndex and yAxisIndex apply only to combo charts.", seriesPath));
                if (xAxisIndex is < 0 or > 1 || yAxisIndex is < 0 or > 1)
                    diagnostics.Add(new("ppj.chart.axisIndexRange", "Combo axis indexes are bounded to primary (0) and secondary (1) axes.", seriesPath));
                if (series.Axis == "primary" && (xAxisIndex == 1 || yAxisIndex == 1))
                    diagnostics.Add(new("ppj.chart.axisIndexConflict", "A primary axis series cannot carry a secondary xAxisIndex or yAxisIndex.", seriesPath));
            }
            if (chart.ChartType != "combo" && series.ChartType is not null)
            {
                var validCandlestickOverlay = chart.ChartType == "candlestick" && index > 0 &&
                    series.ChartType is "line" or "area" or "column";
                if (!validCandlestickOverlay && !string.Equals(series.ChartType, chart.ChartType, StringComparison.Ordinal))
                    diagnostics.Add(new("ppj.chart.seriesType", "A non-combo chart series cannot override the deck chart type except for bounded candlestick overlays.", $"{seriesPath}.chartType"));
            }
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
                if (series.XValues.Count != 0 && !numericCombo)
                    diagnostics.Add(new("ppj.chart.xValueType", "xValues applies only to scatter and bubble charts.", seriesPath + ".xValues"));
                if (series.BubbleSizes.Count != 0)
                    diagnostics.Add(new("ppj.chart.bubbleSizeType", "bubbleSizes applies only to bubble charts.", seriesPath + ".bubbleSizes"));
            }
            if ((chart.ChartType != "candlestick" || index > 0) &&
                (series.OpenValues.Count != 0 || series.HighValues.Count != 0 || series.LowValues.Count != 0))
                diagnostics.Add(new(
                    "ppj.chart.ohlcType",
                    "openValues, highValues, and lowValues apply only to the first candlestick body series.",
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
            if (series.Raw.TryGetProperty("dataLabels", out var dataLabels))
                ValidateSeriesDataLabels(chart, series, dataLabels, seriesPath, diagnostics);
            if (series.Raw.TryGetProperty("pointStyles", out var pointStyles))
                ValidatePointStyles(chart, series, pointStyles, seriesPath, diagnostics);
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
        if (chart.Data.Series.Any(series => series.Raw.TryGetProperty("symbol", out _)))
        {
            if (chart.ChartType is "bar" or "column") ValidatePictographicChart(chart, path, diagnostics);
            else diagnostics.Add(new(
                "ppj.chart.pictographType",
                "Series symbol applies only to bounded bar and column pictographic charts.",
                path + ".data.series"));
        }
        if (chart.ChartType == "combo") ValidateCombo(chart, path, diagnostics);
        if (chart.Raw.TryGetProperty("spokeAxis", out var spokeAxis))
        {
            if (chart.ChartType != "radar")
                diagnostics.Add(new(
                    "ppj.chart.spokeAxisType",
                    "spokeAxis applies only to radar charts.",
                    path + ".spokeAxis"));
            else
                ValidateRadarSpokeAxis(chart, spokeAxis, path, diagnostics);
        }

        if (chart.Raw.TryGetProperty("style", out var style) &&
            style.TryGetProperty("dataLabels", out _) &&
            (style.TryGetProperty("showDataLabels", out _) || style.TryGetProperty("dataLabelPosition", out _)))
            diagnostics.Add(new(
                "ppj.chart.dataLabelConflict",
                "Structured dataLabels cannot be combined with showDataLabels or dataLabelPosition.",
                path + ".style.dataLabels"));
        if (chart.Raw.TryGetProperty("style", out style) &&
            style.TryGetProperty("startAngle", out _) &&
            chart.ChartType is not ("pie" or "doughnut"))
            diagnostics.Add(new(
                "ppj.chart.startAngleType",
                "style.startAngle applies only to pie and doughnut charts.",
                path + ".style.startAngle"));
        if (chart.Raw.TryGetProperty("style", out style) &&
            style.TryGetProperty("holeSize", out _) &&
            chart.ChartType != "doughnut")
            diagnostics.Add(new(
                "ppj.chart.holeSizeType",
                "style.holeSize applies only to doughnut charts.",
                path + ".style.holeSize"));
        if (chart.Raw.TryGetProperty("style", out style) &&
            style.TryGetProperty("varyColors", out _) &&
            chart.ChartType is not ("bar" or "column" or "line" or "combo"))
            diagnostics.Add(new(
                "ppj.chart.varyColorsType",
                "style.varyColors applies only to line, bar, column, and categorical combo charts.",
                path + ".style.varyColors"));
        if (chart.Raw.TryGetProperty("style", out style) &&
            style.TryGetProperty("varyColors", out _) &&
            chart.ChartType == "combo" &&
            !chart.Data.Series.Any(series => series.ChartType is "bar" or "column"))
            diagnostics.Add(new(
                "ppj.chart.varyColorsCombo",
                "style.varyColors on a combo chart requires a bar or column plot.",
                path + ".style.varyColors"));
        var hasBubbleSeries = chart.ChartType == "bubble" ||
            chart.ChartType == "combo" && chart.Data.Series.Any(series => series.ChartType == "bubble");
        if (chart.Raw.TryGetProperty("style", out style) &&
            style.TryGetProperty("dataLabels", out var chartDataLabels) &&
            chartDataLabels.TryGetProperty("showBubbleSize", out var chartBubbleSize) &&
            chartBubbleSize.ValueKind == JsonValueKind.True && !hasBubbleSeries)
            diagnostics.Add(new(
                "ppj.chart.dataLabelBubbleSize",
                "Bubble-size labels require a chart containing a bubble series.",
                path + ".style.dataLabels.showBubbleSize"));
        if (chart.Raw.TryGetProperty("style", out style) &&
            (style.TryGetProperty("bubbleScale", out _) || style.TryGetProperty("bubbleSizeMode", out _) ||
             style.TryGetProperty("bubbleSizeScale", out _) || style.TryGetProperty("bubbleRadiusRange", out _)) &&
            !hasBubbleSeries)
            diagnostics.Add(new(
                "ppj.chart.bubbleStyleType",
                "Bubble sizing style applies only to charts containing a bubble series.",
                path + ".style"));
        if (chart.Raw.TryGetProperty("style", out style) &&
            style.TryGetProperty("bubbleRadiusRange", out var bubbleRadiusRange))
        {
            var radii = bubbleRadiusRange.EnumerateArray().Select(item => item.GetDouble()).ToArray();
            if (radii.Length == 2 && radii[0] >= radii[1])
                diagnostics.Add(new(
                    "ppj.chart.bubbleRadiusRangeOrder",
                    "bubbleRadiusRange requires a strictly increasing [minimum, maximum] radius pair.",
                    path + ".style.bubbleRadiusRange"));
        }
        if (chart.Raw.TryGetProperty("style", out style) &&
            style.TryGetProperty("stacking", out var stacking) &&
            stacking.ValueKind == JsonValueKind.String &&
            stacking.GetString() == "stream")
        {
            if (chart.ChartType != "area")
                diagnostics.Add(new(
                    "ppj.chart.streamType",
                    "style.stacking stream applies only to area charts.",
                    path + ".style.stacking"));
            else ValidateStreamgraph(chart, path, diagnostics);
        }
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
        ValidateAxisKinds(chart.Raw, "xAxis", path, chart.ChartType is not ("scatter" or "bubble") && !numericCombo, diagnostics);
        ValidateAxisKinds(chart.Raw, "yAxis", path, categoryAxis: false, diagnostics);
        ValidateAxisKinds(chart.Raw, "secondaryXAxis", path, categoryAxis: true, diagnostics);
        ValidateAxisKinds(chart.Raw, "secondaryYAxis", path, categoryAxis: false, diagnostics);
    }

    private static void ValidateSeriesDataLabels(
        PpjChartElementModel chart,
        PpjChartSeriesModel series,
        JsonElement labels,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var nativeType = chart.ChartType == "combo" ? series.ChartType : chart.ChartType;
        if (nativeType is not ("column" or "bar" or "line" or "area" or "pie" or "doughnut" or "scatter" or "bubble" or "radar"))
        {
            diagnostics.Add(new(
                "ppj.chart.seriesDataLabelType",
                "Series data-label overrides require a native ChartPart series.",
                path + ".dataLabels"));
            return;
        }

        if (labels.TryGetProperty("showPercent", out var seriesPercent) &&
            seriesPercent.ValueKind == JsonValueKind.True && nativeType is not ("pie" or "doughnut"))
            diagnostics.Add(new(
                "ppj.chart.seriesDataLabelPercent",
                "Percentage labels require a pie or doughnut series.",
                path + ".dataLabels.showPercent"));
        if (labels.TryGetProperty("showBubbleSize", out var seriesBubbleSize) &&
            seriesBubbleSize.ValueKind == JsonValueKind.True && nativeType != "bubble")
            diagnostics.Add(new(
                "ppj.chart.seriesDataLabelBubbleSize",
                "Bubble-size labels require a bubble series.",
                path + ".dataLabels.showBubbleSize"));

        if (!labels.TryGetProperty("points", out var points)) return;
        var previous = -1;
        var pointIndex = 0;
        foreach (var point in points.EnumerateArray())
        {
            var pointPath = $"{path}.dataLabels.points[{pointIndex}]";
            var index = point.GetProperty("index").GetInt32();
            if (index <= previous)
                diagnostics.Add(new(
                    "ppj.chart.dataLabelPointOrder",
                    "Point label indexes must be unique and strictly increasing.",
                    pointPath + ".index"));
            if (index >= series.Values.Count)
                diagnostics.Add(new(
                    "ppj.chart.dataLabelPointRange",
                    "Point label index must address an existing series point.",
                    pointPath + ".index"));
            else if (series.Values[index] is null)
                diagnostics.Add(new(
                    "ppj.chart.dataLabelMissingPoint",
                    "A missing chart point cannot carry a label override.",
                    pointPath + ".index"));
            if (point.TryGetProperty("showPercent", out var pointPercent) &&
                pointPercent.ValueKind == JsonValueKind.True && nativeType is not ("pie" or "doughnut"))
                diagnostics.Add(new(
                    "ppj.chart.dataLabelPointPercent",
                    "Percentage labels require a pie or doughnut series.",
                    pointPath + ".showPercent"));
            if (point.TryGetProperty("showBubbleSize", out var pointBubbleSize) &&
                pointBubbleSize.ValueKind == JsonValueKind.True && nativeType != "bubble")
                diagnostics.Add(new(
                    "ppj.chart.dataLabelPointBubbleSize",
                    "Bubble-size labels require a bubble series.",
                    pointPath + ".showBubbleSize"));
            previous = index;
            pointIndex++;
        }
    }

    private static void ValidatePointStyles(
        PpjChartElementModel chart,
        PpjChartSeriesModel series,
        JsonElement pointStyles,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var nativeType = chart.ChartType == "combo" ? series.ChartType : chart.ChartType;
        if (nativeType is not ("column" or "bar" or "pie" or "doughnut"))
        {
            diagnostics.Add(new(
                "ppj.chart.pointStyleType",
                "Point styles require a native bar, column, pie or doughnut series.",
                path + ".pointStyles"));
            return;
        }

        var previous = -1;
        var itemIndex = 0;
        foreach (var point in pointStyles.EnumerateArray())
        {
            var pointPath = $"{path}.pointStyles[{itemIndex}]";
            var index = point.GetProperty("index").GetInt32();
            if (index <= previous)
                diagnostics.Add(new(
                    "ppj.chart.pointStyleOrder",
                    "Point-style indexes must be unique and strictly increasing.",
                    pointPath + ".index"));
            if (index >= series.Values.Count)
                diagnostics.Add(new(
                    "ppj.chart.pointStyleRange",
                    "Point-style index must address an existing series point.",
                    pointPath + ".index"));
            else if (series.Values[index] is null)
                diagnostics.Add(new(
                    "ppj.chart.pointStyleMissingPoint",
                    "A missing chart point cannot carry a visual override.",
                    pointPath + ".index"));
            if (point.TryGetProperty("explosion", out _) && nativeType is not ("pie" or "doughnut"))
                diagnostics.Add(new(
                    "ppj.chart.pointExplosionType",
                    "Point explosion applies only to pie and doughnut series.",
                    pointPath + ".explosion"));
            previous = index;
            itemIndex++;
        }
    }

    private static void ValidateRadarSpokeAxis(
        PpjChartElementModel chart,
        JsonElement spokeAxis,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        foreach (var axisName in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(axisName, out _))
                diagnostics.Add(new(
                    "ppj.chart.spokeAxisConflict",
                    "Radar spokeAxis cannot be combined with generic or secondary chart axes.",
                    path + "." + axisName));

        if (chart.Raw.TryGetProperty("style", out var style))
            foreach (var propertyName in new[] { "showCategoryAxis", "showValueAxis", "showGridlines" })
                if (style.TryGetProperty(propertyName, out _))
                    diagnostics.Add(new(
                        "ppj.chart.spokeAxisStyleConflict",
                        $"Radar spokeAxis cannot be combined with legacy style.{propertyName}.",
                        path + ".style." + propertyName));

        if (spokeAxis.TryGetProperty("min", out var minimum) &&
            spokeAxis.TryGetProperty("max", out var maximum) &&
            minimum.ValueKind == JsonValueKind.Number &&
            maximum.ValueKind == JsonValueKind.Number &&
            minimum.TryGetDouble(out var minimumValue) &&
            maximum.TryGetDouble(out var maximumValue) &&
            minimumValue >= maximumValue)
            diagnostics.Add(new(
                "ppj.chart.spokeAxisDomain",
                "Radar spokeAxis minimum must be smaller than maximum.",
                path + ".spokeAxis"));
    }

    private static void ValidatePictographicChart(
        PpjChartElementModel chart,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (chart.Data.Series.Count != 1 || !chart.Data.Series[0].Raw.TryGetProperty("symbol", out var symbol))
        {
            diagnostics.Add(new(
                "ppj.chart.pictographSeriesCount",
                "Pictographic bars require exactly one series with a symbol.",
                path + ".data.series"));
            return;
        }
        if (chart.Data.Categories.Count is < 2 or > 12)
            diagnostics.Add(new(
                "ppj.chart.pictographCategories",
                "Pictographic bars require 2..12 categories.",
                path + ".data.categories"));
        if (chart.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            chart.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != chart.Data.Categories.Count)
            diagnostics.Add(new(
                "ppj.chart.pictographCategoryLabel",
                "Pictographic bar categories must be unique non-empty strings.",
                path + ".data.categories"));

        var kind = symbol.GetProperty("kind").GetString();
        if (kind == "icon")
        {
            var iconName = symbol.GetProperty("iconName").GetString()!;
            if (!PpjIconCatalog.Contains(iconName))
                diagnostics.Add(new(
                    "ppj.icon.unknown",
                    $"PPJ iconName {iconName} is not present in the pinned Font Awesome Free catalog.",
                    path + ".data.series[0].symbol.iconName"));
        }
        else
        {
            var preset = symbol.GetProperty("preset").GetString()!;
            if (preset is "textbox" or "line" || !PptxPresetGeometryAdjustmentCodec.HasProfile(preset))
                diagnostics.Add(new(
                    "ppj.chart.pictographPreset",
                    $"Pictographic symbol preset {preset} is not an authored closed-shape preset.",
                    path + ".data.series[0].symbol.preset"));
        }
        if (symbol.TryGetProperty("unitLabel", out var unitLabel) && string.IsNullOrWhiteSpace(unitLabel.GetString()))
            diagnostics.Add(new(
                "ppj.chart.pictographUnitLabel",
                "Pictographic unitLabel must be non-empty when present.",
                path + ".data.series[0].symbol.unitLabel"));

        var unit = symbol.GetProperty("unit").GetDouble();
        var symbolTotal = 0d;
        for (var index = 0; index < chart.Data.Series[0].Values.Count; index++)
        {
            var value = chart.Data.Series[0].Values[index];
            if (value is null || !double.IsFinite(value.Value) || value.Value < 0)
            {
                diagnostics.Add(new(
                    "ppj.chart.pictographValue",
                    "Pictographic values must be finite, complete and non-negative.",
                    $"{path}.data.series[0].values[{index}]"));
                continue;
            }
            var count = value.Value / unit;
            var rounded = Math.Round(count);
            if (Math.Abs(count - rounded) > 1e-9 * Math.Max(1, Math.Abs(count)))
                diagnostics.Add(new(
                    "ppj.chart.pictographUnit",
                    $"Pictographic value {value.Value} is not an exact multiple of unit {unit}.",
                    $"{path}.data.series[0].values[{index}]"));
            if (rounded > 32)
                diagnostics.Add(new(
                    "ppj.chart.pictographCategoryBudget",
                    "A pictographic category may expand to at most 32 symbols.",
                    $"{path}.data.series[0].values[{index}]"));
            if (rounded > 0) symbolTotal += rounded;
        }
        if (symbolTotal > 192)
            diagnostics.Add(new(
                "ppj.chart.pictographTotalBudget",
                "A pictographic chart may expand to at most 192 symbols.",
                path + ".data.series[0].values"));
        if (symbolTotal == 0)
            diagnostics.Add(new(
                "ppj.chart.pictographEmpty",
                "A pictographic chart requires at least one visible symbol.",
                path + ".data.series[0].values"));

        foreach (var property in chart.Data.Series[0].Raw.EnumerateObject())
            if (property.Name is not ("id" or "name" or "values" or "color" or "fill" or "stroke" or "symbol"))
                diagnostics.Add(new(
                    "ppj.chart.pictographSeriesField",
                    $"{property.Name} is not part of the bounded pictographic series profile.",
                    path + ".data.series[0]." + property.Name));
        var series = chart.Data.Series[0].Raw;
        if (series.TryGetProperty("color", out _) && series.TryGetProperty("fill", out _))
            diagnostics.Add(new(
                "ppj.chart.pictographPaint",
                "Pictographic series color and fill are aliases and cannot both be present.",
                path + ".data.series[0]"));
        if (series.TryGetProperty("fill", out var fill) && fill.GetProperty("type").GetString() is not ("solid" or "gradient"))
            diagnostics.Add(new(
                "ppj.chart.pictographPaint",
                "Pictographic symbols support solid or gradient fill.",
                path + ".data.series[0].fill"));
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.pictographAxis",
                    "Pictographic bars generate their own categorical layout and do not accept chart axes.",
                    path + "." + property));
        if (chart.Raw.TryGetProperty("style", out var style))
            foreach (var property in style.EnumerateObject())
                if (property.Name != "titleTextStyle")
                    diagnostics.Add(new(
                        "ppj.chart.pictographStyleField",
                        $"{property.Name} is not part of the bounded pictographic chart style.",
                        path + ".style." + property.Name));
    }

    private static void ValidateStreamgraph(
        PpjChartElementModel chart,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (chart.Data.Categories.Count is < 3 or > 64)
            diagnostics.Add(new(
                "ppj.chart.streamCategories",
                "Streamgraphs require 3..64 ordered categories.",
                path + ".data.categories"));
        if (chart.Data.Categories.Select(category => category.GetRawText()).Distinct(StringComparer.Ordinal).Count() != chart.Data.Categories.Count)
            diagnostics.Add(new(
                "ppj.chart.streamCategoryDuplicate",
                "Streamgraph categories must be unique.",
                path + ".data.categories"));
        if (chart.Data.Series.Count is < 2 or > 12)
            diagnostics.Add(new(
                "ppj.chart.streamSeriesCount",
                "Streamgraphs require 2..12 series.",
                path + ".data.series"));
        if (chart.Data.Series.Any(series => string.IsNullOrWhiteSpace(series.Name)) ||
            chart.Data.Series.Select(series => series.Name).Distinct(StringComparer.Ordinal).Count() != chart.Data.Series.Count)
            diagnostics.Add(new(
                "ppj.chart.streamSeriesName",
                "Streamgraph series names must be unique and non-empty.",
                path + ".data.series"));
        if (chart.Data.Series.SelectMany(series => series.Values).Any(value => value is null or < 0))
            diagnostics.Add(new(
                "ppj.chart.streamValue",
                "Streamgraph values must be finite and non-negative.",
                path + ".data.series"));
        for (var categoryIndex = 0; categoryIndex < chart.Data.Categories.Count; categoryIndex++)
            if (chart.Data.Series.Sum(series => series.Values.ElementAtOrDefault(categoryIndex) ?? 0) <= 0)
                diagnostics.Add(new(
                    "ppj.chart.streamEmptyCategory",
                    "Every streamgraph category requires a positive total.",
                    $"{path}.data.categories[{categoryIndex}]"));
        for (var seriesIndex = 0; seriesIndex < chart.Data.Series.Count; seriesIndex++)
        {
            var series = chart.Data.Series[seriesIndex];
            foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "parents", "sources", "targets", "chartType", "axis", "marker", "trendlines", "errorBars" })
                if (series.Raw.TryGetProperty(property, out _))
                    diagnostics.Add(new(
                        "ppj.chart.streamSeriesField",
                        $"{property} is not part of the bounded streamgraph series profile.",
                        $"{path}.data.series[{seriesIndex}].{property}"));
        }
        foreach (var property in new[] { "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.streamAxis",
                    "Streamgraphs use one generated centered value scale and do not accept Y or secondary axes.",
                    path + "." + property));
        if (chart.Raw.TryGetProperty("style", out var style))
        {
            foreach (var property in style.EnumerateObject())
                if (property.Name is not ("stacking" or "legend" or "legendOverlay" or "titleTextStyle" or "legendTextStyle"))
                    diagnostics.Add(new(
                        "ppj.chart.streamStyleField",
                        $"{property.Name} is not part of the bounded streamgraph style profile.",
                        path + ".style." + property.Name));
            if (style.TryGetProperty("legendOverlay", out _))
                diagnostics.Add(new(
                    "ppj.chart.streamLegendOverlay",
                    "Streamgraph legends do not support overlay.",
                    path + ".style.legendOverlay"));
        }
    }

    private static readonly string[] VectorChartNumberFormats = ["0", "0.0", "0.00", "#,##0", "#,##0.0", "#,##0.00"];

    private static bool IsNumericCombo(PpjChartElementModel chart) =>
        chart.ChartType == "combo" &&
        chart.Data.Series.Any(series => series.ChartType is "scatter" or "bubble");

    private static void ValidateCombo(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        if (IsNumericCombo(chart))
        {
            ValidateNumericCombo(chart, path, diagnostics);
            return;
        }

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

    private static void ValidateNumericCombo(PpjChartElementModel chart, string path, List<PpjDiagnostic> diagnostics)
    {
        if (chart.Data.Categories.Count != 0)
            diagnostics.Add(new(
                "ppj.chart.numericComboCategories",
                "Numeric combo charts require an empty shared categories array.",
                path + ".data.categories"));
        if (chart.Data.Series.Count is < 2 or > 8)
            diagnostics.Add(new(
                "ppj.chart.numericComboSeriesCount",
                "Numeric combo charts require 2..8 series.",
                path + ".data.series"));

        var families = chart.Data.Series.Select(series => series.ChartType).Distinct(StringComparer.Ordinal).ToArray();
        if (!families.Any(family => family is "scatter" or "bubble") || families.Length < 2)
            diagnostics.Add(new(
                "ppj.chart.numericComboFamilies",
                "Numeric combo charts require scatter or bubble evidence and at least one different plot family.",
                path + ".data.series"));

        for (var index = 0; index < chart.Data.Series.Count; index++)
        {
            var series = chart.Data.Series[index];
            var seriesPath = $"{path}.data.series[{index}]";
            if (series.ChartType is not ("scatter" or "bubble" or "line" or "area" or "column"))
                diagnostics.Add(new(
                    "ppj.chart.numericComboType",
                    "Numeric combo series types are scatter, bubble, line, area and column.",
                    seriesPath + ".chartType"));
            if (series.Values.Count is < 2 or > 64 || series.Values.Any(value => value is null || !double.IsFinite(value.Value)))
                diagnostics.Add(new(
                    "ppj.chart.numericComboValues",
                    "Every numeric combo series requires 2..64 complete finite values.",
                    seriesPath + ".values"));
            if (series.XValues.Count != series.Values.Count || series.XValues.Any(value => !double.IsFinite(value)))
                diagnostics.Add(new(
                    "ppj.chart.numericComboXValues",
                    "Every numeric combo series requires one finite xValue per value.",
                    seriesPath + ".xValues"));
            else if (series.XValues.Zip(series.XValues.Skip(1)).Any(pair => pair.First >= pair.Second))
                diagnostics.Add(new(
                    "ppj.chart.numericComboXOrder",
                    "Numeric combo xValues must be strictly increasing within each series.",
                    seriesPath + ".xValues"));
            if (series.ChartType == "bubble")
            {
                if (series.BubbleSizes.Count != series.Values.Count || series.BubbleSizes.Any(value => !double.IsFinite(value) || value <= 0))
                    diagnostics.Add(new(
                        "ppj.chart.numericComboBubbleSizes",
                        "Bubble series require one finite positive bubbleSize per value.",
                        seriesPath + ".bubbleSizes"));
            }
            else if (series.BubbleSizes.Count != 0)
                diagnostics.Add(new(
                    "ppj.chart.numericComboBubbleType",
                    "bubbleSizes apply only to a bubble series.",
                    seriesPath + ".bubbleSizes"));
            if (series.Raw.TryGetProperty("marker", out var marker))
            {
                if (series.ChartType is "area" or "column" or "bubble")
                    diagnostics.Add(new(
                        "ppj.chart.numericComboMarkerType",
                        $"Markers are not rendered for numeric combo {series.ChartType} series.",
                        seriesPath + ".marker"));
                else if (series.ChartType == "scatter" &&
                         (marker.ValueKind == JsonValueKind.String && marker.GetString() == "none" ||
                          marker.ValueKind == JsonValueKind.Object && marker.TryGetProperty("symbol", out var symbol) && symbol.GetString() == "none"))
                    diagnostics.Add(new(
                        "ppj.chart.numericComboScatterMarker",
                        "Scatter series cannot use marker none because that would make the series invisible.",
                        seriesPath + ".marker"));
            }
            if (series.Axis is not null and not "primary")
                diagnostics.Add(new(
                    "ppj.chart.numericComboAxis",
                    "The bounded numeric combo profile uses one shared primary value/value axis pair.",
                    seriesPath + ".axis"));
            foreach (var property in new[] { "pointRoles", "openValues", "highValues", "lowValues", "parents", "sources", "targets", "symbol", "trendlines", "errorBars" })
                if (series.Raw.TryGetProperty(property, out _))
                    diagnostics.Add(new(
                        "ppj.chart.numericComboSeriesField",
                        $"{property} is not part of the bounded numeric combo profile.",
                        $"{seriesPath}.{property}"));
        }

        foreach (var property in new[] { "secondaryXAxis", "secondaryYAxis" })
            if (chart.Raw.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.numericComboSecondaryAxis",
                    "Numeric combo charts do not support secondary axes.",
                    path + "." + property));

        if (chart.Raw.TryGetProperty("style", out var style))
            foreach (var property in style.EnumerateObject())
                if (property.Name is not ("legend" or "legendOverlay" or "titleTextStyle" or "legendTextStyle" or "bubbleScale" or "bubbleSizeMode" or "bubbleSizeScale" or "bubbleRadiusRange"))
                    diagnostics.Add(new(
                        "ppj.chart.numericComboStyleField",
                        $"{property.Name} is not part of the bounded numeric combo style profile.",
                        path + ".style." + property.Name));
        if (style.ValueKind == JsonValueKind.Object && style.TryGetProperty("legend", out var legend) &&
            legend.ValueKind == JsonValueKind.String && legend.GetString() is not ("none" or "right"))
            diagnostics.Add(new(
                "ppj.chart.numericComboLegend",
                "Numeric combo legends support only none or right.",
                path + ".style.legend"));
        if (style.ValueKind == JsonValueKind.Object && style.TryGetProperty("legendOverlay", out _))
            diagnostics.Add(new(
                "ppj.chart.numericComboLegendOverlay",
                "Numeric combo legends do not support overlay.",
                path + ".style.legendOverlay"));
        if (style.ValueKind == JsonValueKind.Object && style.TryGetProperty("bubbleScale", out var bubbleScale) &&
            bubbleScale.ValueKind == JsonValueKind.Number && bubbleScale.GetInt32() < 10)
            diagnostics.Add(new(
                "ppj.chart.numericComboBubbleScale",
                "Generated numeric combo bubbles require bubbleScale between 10 and 300.",
                path + ".style.bubbleScale"));

        ValidateNumericVectorAxis(chart.Raw, "xAxis", path, diagnostics);
        ValidateNumericVectorAxis(chart.Raw, "yAxis", path, diagnostics);
        var requiresZero = chart.Data.Series.Any(series => series.ChartType is "area" or "column");
        if (requiresZero && chart.Raw.TryGetProperty("yAxis", out var yAxis) &&
            (yAxis.TryGetProperty("min", out var minimum) && minimum.ValueKind == JsonValueKind.Number && minimum.GetDouble() > 0 ||
             yAxis.TryGetProperty("max", out var maximum) && maximum.ValueKind == JsonValueKind.Number && maximum.GetDouble() < 0))
            diagnostics.Add(new(
                "ppj.chart.numericComboBaseline",
                "Area and column overlays require zero inside the explicit Y domain.",
                path + ".yAxis"));
    }

    private static void ValidateNumericVectorAxis(
        JsonElement chart,
        string axisName,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (!chart.TryGetProperty(axisName, out var axis)) return;
        foreach (var property in axis.EnumerateObject())
            if (property.Name is not ("visible" or "title" or "numberFormat" or "min" or "max" or "majorUnit" or "textStyle" or "titleTextStyle" or "reverse" or "axisLine" or "gridLine"))
                diagnostics.Add(new(
                    "ppj.chart.numericComboAxisField",
                    $"{property.Name} is not supported by a generated numeric axis.",
                    $"{path}.{axisName}.{property.Name}"));
        if (axis.TryGetProperty("numberFormat", out var numberFormat))
        {
            if (numberFormat.ValueKind == JsonValueKind.Object && numberFormat.TryGetProperty("token", out _))
                diagnostics.Add(new(
                    "ppj.chart.numericComboNumberFormat",
                    "Generated numeric axes require a literal number format; grammar tokens are supported only by native ChartPart axes.",
                    $"{path}.{axisName}.numberFormat"));
            else if (numberFormat.ValueKind != JsonValueKind.String ||
                     !VectorChartNumberFormats.Contains(numberFormat.GetString()!, StringComparer.Ordinal))
                diagnostics.Add(new(
                    "ppj.chart.numericComboNumberFormat",
                    $"Generated numeric axes support {string.Join(", ", VectorChartNumberFormats)}.",
                    $"{path}.{axisName}.numberFormat"));
        }
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
        foreach (var property in new[] { "legend", "legendOverlay", "stacking", "gapWidth", "overlap", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall" })
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
        if (chart.Data.Series.Count is < 1 or > 5)
        {
            diagnostics.Add(new(
                "ppj.chart.candlestickSeriesCount",
                "Candlestick charts require one OHLC or HLC body and at most four overlay series.",
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

            var overlayValues = chart.Data.Series.Skip(1).SelectMany(item => item.Values).Where(value => value is not null).Select(value => value!.Value);
            var lowest = Math.Min(series.LowValues.Min(), overlayValues.DefaultIfEmpty(double.PositiveInfinity).Min());
            var highest = Math.Max(series.HighValues.Max(), overlayValues.DefaultIfEmpty(double.NegativeInfinity).Max());
            if (chart.Data.Series.Skip(1).Any(item => item.ChartType is "area" or "column"))
            {
                lowest = Math.Min(lowest, 0);
                highest = Math.Max(highest, 0);
            }
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

        for (var index = 1; index < chart.Data.Series.Count; index++)
        {
            var overlay = chart.Data.Series[index];
            var overlayPath = $"{path}.data.series[{index}]";
            if (overlay.ChartType is not ("line" or "area" or "column"))
                diagnostics.Add(new(
                    "ppj.chart.candlestickOverlayType",
                    "Candlestick overlays support line, area and column series.",
                    overlayPath + ".chartType"));
            if (overlay.Values.Count != count || overlay.Values.Any(value => value is null || !double.IsFinite(value.Value)))
                diagnostics.Add(new(
                    "ppj.chart.candlestickOverlayValues",
                    "Candlestick overlays require one complete finite value per category.",
                    overlayPath + ".values"));
            foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "parents", "sources", "targets", "axis", "symbol", "trendlines", "errorBars" })
                if (overlay.Raw.TryGetProperty(property, out _))
                    diagnostics.Add(new(
                        "ppj.chart.candlestickOverlayField",
                        $"{property} is not part of the bounded candlestick overlay profile.",
                        $"{overlayPath}.{property}"));
            if (overlay.Raw.TryGetProperty("fill", out _) && overlay.Raw.TryGetProperty("color", out _))
                diagnostics.Add(new(
                    "ppj.chart.candlestickOverlayPaint",
                    "Candlestick overlay color and fill are aliases and cannot both be present.",
                    overlayPath));
            if (overlay.ChartType is "area" or "column" && overlay.Raw.TryGetProperty("marker", out _))
                diagnostics.Add(new(
                    "ppj.chart.candlestickOverlayMarker",
                    $"Markers are not rendered for candlestick {overlay.ChartType} overlays.",
                    overlayPath + ".marker"));
        }

        if (!chart.Raw.TryGetProperty("style", out var style) || !style.TryGetProperty("candlestick", out var candlestick)) return;
        foreach (var property in new[] { "legend", "legendOverlay", "stacking", "gapWidth", "overlap", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap" })
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
        if (series.Levels is < 1 or > 8)
            diagnostics.Add(new(
                "ppj.chart.treemapLevels",
                "Treemap display levels must be between one and eight.",
                seriesPath + ".levels"));
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
        foreach (var property in new[] { "legend", "legendOverlay", "stacking", "gapWidth", "overlap", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick" })
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
        if (series.Levels is < 1 or > 6)
            diagnostics.Add(new(
                "ppj.chart.sunburstLevels",
                "Sunburst display levels must be between one and six.",
                seriesPath + ".levels"));
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
        foreach (var property in new[] { "legend", "legendOverlay", "stacking", "gapWidth", "overlap", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap" })
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

        if (!chart.Raw.TryGetProperty("style", out var style) || !style.TryGetProperty("sankey", out var sankey)) return;
        if (sankey.TryGetProperty("nodeColorMap", out var nodeColorMap))
            foreach (var property in nodeColorMap.EnumerateObject())
                if (!nodes.Contains(property.Name))
                    diagnostics.Add(new(
                        "ppj.chart.sankeyNodeColor",
                        $"Sankey nodeColorMap key {property.Name} is not a declared node.",
                        $"{path}.style.sankey.nodeColorMap.{property.Name}"));
        foreach (var property in new[] { "legend", "legendOverlay", "stacking", "gapWidth", "overlap", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap", "sunburst" })
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
        foreach (var property in new[] { "stacking", "showDataLabels", "dataLabelPosition", "dataLabels", "smooth", "varyColors", "legendOverlay" })
            if (style.TryGetProperty(property, out _))
                diagnostics.Add(new(
                    "ppj.chart.waterfallStyleField",
                    $"{property} is not part of the bounded waterfall style profile.",
                    $"{path}.style.{property}"));
        if (style.TryGetProperty("legend", out var legend) &&
            legend.ValueKind == JsonValueKind.String && legend.GetString() != "none")
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
        if (axis.TryGetProperty("axisLineArrow", out _) &&
            axis.TryGetProperty("axisLine", out var axisLine) &&
            axisLine.ValueKind == JsonValueKind.False)
            diagnostics.Add(new(
                "ppj.chart.axisArrowHiddenLine",
                "axisLineArrow requires a visible axis line.",
                $"{path}.{property}.axisLineArrow"));
        if (axis.TryGetProperty("tickLabelPosition", out var tickLabelPosition) &&
            tickLabelPosition.ValueKind == JsonValueKind.String &&
            tickLabelPosition.GetString() is not ("nextTo" or "high" or "low" or "none"))
            diagnostics.Add(new(
                "ppj.chart.axisTickLabelPosition",
                "tickLabelPosition must be nextTo, high, low, or none.",
                $"{path}.{property}.tickLabelPosition"));
        if (axis.TryGetProperty("majorTickMark", out var majorTickMark) &&
            majorTickMark.ValueKind == JsonValueKind.String &&
            majorTickMark.GetString() is not ("cross" or "in" or "out" or "none"))
            diagnostics.Add(new(
                "ppj.chart.axisMajorTickMark",
                "majorTickMark must be cross, in, out, or none.",
                $"{path}.{property}.majorTickMark"));
        if (axis.TryGetProperty("minorTickMark", out var minorTickMark) &&
            minorTickMark.ValueKind == JsonValueKind.String &&
            minorTickMark.GetString() is not ("cross" or "in" or "out" or "none"))
            diagnostics.Add(new(
                "ppj.chart.axisMinorTickMark",
                "minorTickMark must be cross, in, out, or none.",
                $"{path}.{property}.minorTickMark"));
        if (axis.TryGetProperty("position", out var position) &&
            position.ValueKind == JsonValueKind.String &&
            position.GetString() is not ("bottom" or "left" or "right" or "top"))
            diagnostics.Add(new(
                "ppj.chart.axisPosition",
                "position must be bottom, left, right, or top.",
                $"{path}.{property}.position"));
        if (axis.TryGetProperty("tickLabelsVisible", out var tickLabelsVisible) &&
            axis.TryGetProperty("tickLabelPosition", out tickLabelPosition) &&
            tickLabelsVisible.ValueKind == JsonValueKind.False &&
            tickLabelPosition.ValueKind == JsonValueKind.String &&
            tickLabelPosition.GetString() != "none")
            diagnostics.Add(new(
                "ppj.chart.axisTickLabelConflict",
                "tickLabelsVisible=false may accompany tickLabelPosition only when the position is none.",
                $"{path}.{property}"));
        if (axis.TryGetProperty("tickLabelsVisible", out tickLabelsVisible) &&
            axis.TryGetProperty("tickLabelPosition", out tickLabelPosition) &&
            tickLabelsVisible.ValueKind == JsonValueKind.True &&
            tickLabelPosition.ValueKind == JsonValueKind.String)
            diagnostics.Add(new(
                "ppj.chart.axisTickLabelConflict",
                "tickLabelsVisible=true cannot accompany an explicit tickLabelPosition.",
                $"{path}.{property}"));
        if (categoryAxis)
        {
            foreach (var name in new[] { "min", "max", "majorUnit", "minorUnit" })
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
        // Grammar token references are resolved by the authored/source-bound
        // compiler after schema and token-kind validation.  The semantic
        // validator cannot compare their values without duplicating the
        // compiler's precedence rules, so it only performs this early check
        // for literal numeric bounds.
        if (minimum.ValueKind == JsonValueKind.Number &&
            maximum.ValueKind == JsonValueKind.Number &&
            minimum.TryGetDouble(out var minimumValue) &&
            maximum.TryGetDouble(out var maximumValue) &&
            minimumValue >= maximumValue)
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
        IReadOnlyList<PpjAssetModel> assets,
        List<PpjDiagnostic> diagnostics)
    {
        if (smartArt.Mode == "authored" && (smartArt.Layout is null) == (smartArt.DefinitionAssetId is null))
            diagnostics.Add(new(
                "ppj.smartArt.definition",
                "Authored SmartArt requires exactly one built-in layout or definitionAsset.",
                path));
        if (smartArt.Mode == "source-bound" && smartArt.NativeRef is null && smartArt.Nodes.All(node => node.NativeRef is null))
            diagnostics.Add(new("ppj.smartArt.nativeRef", "Source-bound SmartArt requires an element or node nativeRef.", $"{path}.nativeRef"));
        if (smartArt.DefinitionAssetId is not null)
        {
            ValidateAssetRef(smartArt.DefinitionAssetId, assetIds, $"{path}.definitionAsset", diagnostics);
            var definition = assets.FirstOrDefault(asset => asset.Id == smartArt.DefinitionAssetId);
            if (definition is not null && !definition.MimeType.Equals("application/vnd.officekit.smartart-definition+json", StringComparison.OrdinalIgnoreCase))
                diagnostics.Add(new(
                    "ppj.smartArt.definitionMime",
                    "SmartArt definitionAsset requires application/vnd.officekit.smartart-definition+json.",
                    $"{path}.definitionAsset"));
        }

        if (smartArt.Mode == "authored")
        {
            if (smartArt.DetachToShapes)
                diagnostics.Add(new(
                    "ppj.smartArt.detachSourceOnly",
                    "detachToShapes is available only for a source-bound SmartArt with a verified cached drawing.",
                    $"{path}.detachToShapes"));
            if (smartArt.LayoutDefinitionId is not null)
                diagnostics.Add(new(
                    "ppj.smartArt.sourceIdentity",
                    "Authored diagrams cannot declare a source layoutDefinitionId.",
                    $"{path}.layoutDefinitionId"));
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

            var hasParentEdges = smartArt.Connections.Any(connection => connection.Role == "parent");
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
        var connections = UniqueIndex(smartArt.Connections, connection => connection.Id, $"{path}.connections", diagnostics);
        _ = connections;
        var explicitParents = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 0; index < smartArt.Connections.Count; index++)
        {
            var connection = smartArt.Connections[index];
            var connectionPath = $"{path}.connections[{index}]";
            if (!nodes.ContainsKey(connection.FromId))
                diagnostics.Add(new("ppj.smartArt.connectionEndpoint", $"SmartArt connection source {connection.FromId} does not exist.", $"{connectionPath}.from"));
            if (!nodes.ContainsKey(connection.ToId))
                diagnostics.Add(new("ppj.smartArt.connectionEndpoint", $"SmartArt connection destination {connection.ToId} does not exist.", $"{connectionPath}.to"));
            if (connection.FromId == connection.ToId)
                diagnostics.Add(new("ppj.smartArt.connectionLoop", "SmartArt connections cannot target their own source node.", connectionPath));
            if (connection.Role == "parent" && !explicitParents.TryAdd(connection.ToId, connection.FromId))
                diagnostics.Add(new("ppj.smartArt.parent", $"SmartArt node {connection.ToId} has more than one parent connection.", connectionPath));
        }
        var rawNodes = smartArt.Raw.GetProperty("nodes").EnumerateArray().ToArray();
        for (var index = 0; index < smartArt.Nodes.Count; index++)
        {
            var node = smartArt.Nodes[index];
            var nodePath = $"{path}.nodes[{index}]";
            if (smartArt.Mode == "authored" && rawNodes[index].TryGetProperty("kind", out _))
                diagnostics.Add(new(
                    "ppj.smartArt.sourceIdentity",
                    "Authored SmartArt node kind is derived by the native writer and cannot be declared.",
                    $"{nodePath}.kind"));
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
            if (node.Text.Paragraphs.SelectMany(paragraph => paragraph.Runs).Any(run => run.Formula is not null))
                diagnostics.Add(new("ppj.smartArt.formula", "Authored SmartArt node text does not support formula runs.", $"{nodePath}.text"));
            ValidateNativeRef(node.NativeRef, source, $"{nodePath}.nativeRef", diagnostics);
        }

        var parents = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (childId, parentId) in explicitParents) parents[childId] = parentId;
        foreach (var node in smartArt.Nodes)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var cursor = node.Id;
            while (parents.TryGetValue(cursor, out var parentId) && nodes.ContainsKey(parentId))
            {
                if (!seen.Add(cursor))
                {
                    diagnostics.Add(new("ppj.smartArt.cycle", "SmartArt parent references form a cycle.", $"{path}.nodes"));
                    break;
                }
                cursor = parentId;
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
        PpjProgramModel program,
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
                var suppliedElement = supplied.Value[index];
                if (!slot.Accepts.Contains(suppliedElement.Type, StringComparer.Ordinal))
                    diagnostics.Add(new("ppj.component.slotType", $"Slot {supplied.Key} does not accept {supplied.Value[index].Type} elements.", $"{path}.slots.{supplied.Key}[{index}].type"));
                if (slot.ImagePolicy is not null)
                {
                    if (suppliedElement is not PpjImageElementModel image)
                    {
                        diagnostics.Add(new(
                            "ppj.component.imagePolicyType",
                            $"Slot {supplied.Key} image policy only accepts image elements.",
                            $"{path}.slots.{supplied.Key}[{index}]"));
                    }
                    else
                    {
                        ValidateImageSlotPolicy(image, slot.ImagePolicy, program.Root, program.Assets, $"{path}.slots.{supplied.Key}[{index}]", diagnostics);
                    }
                }
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
            var layoutDirection = instance.Repeat.Direction ?? "vertical";
            if (layoutDirection is not ("horizontal" or "vertical" or "grid" or "flow"))
                diagnostics.Add(new("ppj.component.repeatLayout", "Repeat layout direction must be horizontal, vertical, grid, or flow.", $"{path}.repeat.layout.direction"));
            if (layoutDirection == "grid" && instance.Repeat.Columns is null)
                diagnostics.Add(new("ppj.component.repeatGridColumns", "Grid repeat layout requires an explicit positive columns value.", $"{path}.repeat.layout.columns"));
            if (layoutDirection == "flow" && instance.Repeat.Columns is not null && instance.Repeat.Columns < 1)
                diagnostics.Add(new("ppj.component.repeatFlowColumns", "Flow repeat layout columns must be positive.", $"{path}.repeat.layout.columns"));
            if (layoutDirection is not ("grid" or "flow") && instance.Repeat.Columns is not null)
                diagnostics.Add(new("ppj.component.repeatGridColumns", "Repeat layout columns is only valid for grid or flow direction.", $"{path}.repeat.layout.columns"));
            if (instance.Repeat.Anchor is not (null or "start" or "center" or "end"))
                diagnostics.Add(new("ppj.component.repeatAnchor", "Repeat layout anchor must be start, center, or end.", $"{path}.repeat.layout.anchor"));
            if (layoutDirection is not ("grid" or "flow") && instance.Repeat.RowGap is not null)
                diagnostics.Add(new("ppj.component.repeatRowGap", "Repeat layout rowGap is only valid for grid or flow direction.", $"{path}.repeat.layout.rowGap"));
            if (instance.Repeat.Weights is not null)
            {
                if (layoutDirection is not ("horizontal" or "vertical"))
                    diagnostics.Add(new("ppj.component.repeatStackWeightsDirection", "Repeat layout weights are only valid for horizontal or vertical direction.", $"{path}.repeat.layout.weights"));
                if (instance.Repeat.Anchor is not null)
                    diagnostics.Add(new("ppj.component.repeatStackAnchor", "Repeat layout anchor cannot be combined with weighted stack allocation.", $"{path}.repeat.layout.anchor"));
                if (instance.Repeat.Weights.Count != instance.Repeat.Items.Count)
                    diagnostics.Add(new("ppj.component.repeatStackWeightsCount", "Repeat layout weights must contain exactly one value per repeat item.", $"{path}.repeat.layout.weights"));
                for (var index = 0; index < instance.Repeat.Weights.Count; index++)
                {
                    var weight = instance.Repeat.Weights[index];
                    if (!double.IsFinite(weight) || weight <= 0 || weight > 100000)
                        diagnostics.Add(new("ppj.component.repeatStackWeight", "Repeat stack weights must be finite, positive, and no greater than 100000.", $"{path}.repeat.layout.weights[{index}]"));
                }
            }
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

    private static void ValidateImageSlotPolicy(
        PpjImageElementModel image,
        PpjImageSlotPolicyModel policy,
        JsonElement root,
        IReadOnlyList<PpjAssetModel> assets,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        var fit = ResolveImageFit(root, image.Raw, image.Fit, path, diagnostics) ?? "contain";
        if (policy.AllowedFit.Count > 0 && !policy.AllowedFit.Contains(fit, StringComparer.Ordinal))
            diagnostics.Add(new("ppj.component.imageFit", $"Image fit {fit} is not allowed by slot policy {policy.Role}.", $"{path}.fit"));
        var mask = image.MaskKind switch
        {
            "preset" => image.MaskPreset ?? "none",
            null => "none",
            var value => value,
        };
        if (policy.AllowedMask.Count > 0 && !policy.AllowedMask.Contains(mask, StringComparer.Ordinal))
            diagnostics.Add(new("ppj.component.imageMask", $"Image mask {mask} is not allowed by slot policy {policy.Role}.", $"{path}.mask"));
        var asset = assets.FirstOrDefault(candidate => candidate.Id.Equals(image.AssetId, StringComparison.Ordinal));
        if (asset is null) return;
        if (policy.MinimumWidthPx is { } minWidth && (!asset.WidthPx.HasValue || asset.WidthPx.Value < minWidth))
            diagnostics.Add(new("ppj.component.imageDimensions", $"Image asset {image.AssetId} is narrower than the slot minimum of {minWidth}px.", $"{path}.asset"));
        if (policy.MinimumHeightPx is { } minHeight && (!asset.HeightPx.HasValue || asset.HeightPx.Value < minHeight))
            diagnostics.Add(new("ppj.component.imageDimensions", $"Image asset {image.AssetId} is shorter than the slot minimum of {minHeight}px.", $"{path}.asset"));
        if (policy.Rights.Count > 0)
        {
            var status = asset.Rights.ValueKind == JsonValueKind.Object && asset.Rights.TryGetProperty("status", out var statusValue)
                ? statusValue.GetString()
                : null;
            if (status is null || !policy.Rights.Contains(status, StringComparer.Ordinal))
                diagnostics.Add(new("ppj.component.imageRights", $"Image asset {image.AssetId} has rights status {status ?? "(missing)"}, outside slot policy {policy.Role}.", $"{path}.asset"));
        }
    }

    private static void ValidateImageFocus(
        JsonElement raw,
        string? fit,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (!raw.TryGetProperty("focus", out var focus)) return;
        if (raw.TryGetProperty("crop", out _))
            diagnostics.Add(new(
                "ppj.image.focusCrop",
                "An image cannot declare both an explicit crop and a focal crop.",
                path + ".focus"));
        if (!string.Equals(fit, "cover", StringComparison.Ordinal))
            diagnostics.Add(new(
                "ppj.image.focusFit",
                "An image focal point requires cover fit so the compiler can derive an asymmetric crop.",
                path + ".focus"));
        if (focus.ValueKind != JsonValueKind.Object ||
            !focus.TryGetProperty("x", out var x) ||
            !focus.TryGetProperty("y", out var y) ||
            !x.TryGetDouble(out var xValue) ||
            !y.TryGetDouble(out var yValue) ||
            !double.IsFinite(xValue) ||
            !double.IsFinite(yValue) ||
            xValue is < 0 or > 1 ||
            yValue is < 0 or > 1)
        {
            diagnostics.Add(new(
                "ppj.image.focus",
                "Image focal point x and y must be finite normalized values between 0 and 1.",
                path + ".focus"));
        }
    }

    private static string? ResolveImageFit(
        JsonElement root,
        JsonElement raw,
        string? literalFit,
        string path,
        List<PpjDiagnostic> diagnostics)
    {
        if (literalFit is not null || !raw.TryGetProperty("fit", out var fit)) return literalFit;
        if (fit.ValueKind != JsonValueKind.Object ||
            !fit.TryGetProperty("token", out var tokenValue) ||
            tokenValue.ValueKind != JsonValueKind.String)
            return literalFit;
        var token = tokenValue.GetString()!;
        var tokens = root.GetProperty("design").TryGetProperty("grammar", out var grammar) &&
            grammar.TryGetProperty("tokens", out var tokenMap) &&
            tokenMap.ValueKind == JsonValueKind.Object
            ? tokenMap
            : default;
        if (tokens.ValueKind != JsonValueKind.Object || !tokens.TryGetProperty(token, out var definition))
        {
            diagnostics.Add(new("ppj.grammar.tokenUnknown", $"PPJ grammar token {token} for image fit is not declared.", $"{path}.fit"));
            return null;
        }
        if (definition.ValueKind != JsonValueKind.Object ||
            !definition.TryGetProperty("kind", out var kind) ||
            !string.Equals(kind.GetString(), "string", StringComparison.Ordinal))
        {
            diagnostics.Add(new("ppj.grammar.tokenKind", $"PPJ grammar token {token} for image fit must declare kind string.", $"{path}.fit"));
            return null;
        }
        if (!definition.TryGetProperty("value", out var value) || value.ValueKind != JsonValueKind.String)
        {
            diagnostics.Add(new("ppj.grammar.tokenValue", $"PPJ grammar token {token} for image fit must resolve to a string.", $"{path}.fit"));
            return null;
        }
        var resolved = value.GetString();
        if (resolved is not ("cover" or "contain" or "stretch" or "tile" or "none"))
            diagnostics.Add(new("ppj.image.fit", $"Image fit {resolved ?? "(empty)"} is outside the bounded profile.", $"{path}.fit"));
        return resolved;
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
            "crop" => IsWrappedValue(value, "crop"),
            "focus" => IsWrappedValue(value, "focus"),
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
            if (animation.ChartBuild is not null && target is PpjChartElementModel streamChart && IsInlineStreamgraph(streamChart))
                diagnostics.Add(new(
                    "ppj.animation.vectorChartBuild",
                    "Vector-lowered streamgraphs support whole-object animation, not ChartPart build modes.",
                    $"{path}.chartBuild"));
            if (animation.ChartBuild is not null && target is PpjChartElementModel pictographChart && IsInlinePictographicChart(pictographChart))
                diagnostics.Add(new(
                    "ppj.animation.vectorChartBuild",
                    "Vector-lowered pictographic bars support whole-object animation, not ChartPart build modes.",
                    $"{path}.chartBuild"));
            if (animation.ChartBuild is not null && target is PpjChartElementModel numericCombo && IsNumericCombo(numericCombo))
                diagnostics.Add(new(
                    "ppj.animation.vectorChartBuild",
                    "Vector-lowered numeric combo charts support whole-object animation, not ChartPart build modes.",
                    $"{path}.chartBuild"));
            if (animation.ChartBuild is not null && target is PpjChartElementModel sizedBubble && IsInlineSizedBubble(sizedBubble))
                diagnostics.Add(new(
                    "ppj.animation.vectorChartBuild",
                    "Explicitly sized bubble charts support whole-object animation, not ChartPart build modes.",
                    $"{path}.chartBuild"));
            if ((animation.Effect == "pulse") != (animation.Phase == "emphasis"))
                diagnostics.Add(new("ppj.animation.phaseEffect", "pulse is the only emphasis effect and is only valid in the emphasis phase.", $"{path}.effect"));
            if (animation.Repeat is < 1 or > 8)
                diagnostics.Add(new("ppj.animation.repeat", "Animation repeat must be between 1 and 8.", $"{path}.repeat"));
            if (animation.Easing is not (null or "linear" or "ease-in" or "ease-out" or "ease-in-out"))
                diagnostics.Add(new("ppj.animation.easing", "Animation easing is outside the bounded native profile.", $"{path}.easing"));

            expandedTimingNodes += EstimateTimingNodes(animation, target);
        }
        if (expandedTimingNodes > 64)
            diagnostics.Add(new("ppj.animation.timingBudget", $"Page expands to {expandedTimingNodes} timing nodes; the limit is 64.", $"{pagePath}.animations"));

        // The compact animation array accepts the same explicit trigger sugar
        // as timingGraph.  Keep `start` authoritative for `timeline`, but do
        // not let a mismatched trigger silently change the native schedule.
        if (page.Raw.TryGetProperty("animations", out var rawAnimations))
        {
            for (var index = 0; index < rawAnimations.GetArrayLength(); index++)
            {
                var animation = rawAnimations[index];
                var path = $"{pagePath}.animations[{index}]";
                if (!animation.TryGetProperty("trigger", out var trigger)) continue;
                var triggerValue = trigger.GetString();
                if (triggerValue is not (null or "timeline" or "onClick" or "afterPrevious" or "withPrevious"))
                    diagnostics.Add(new("ppj.animation.triggerUnsupported", "The animation trigger is outside the bounded click/previous profile.", path + ".trigger"));
                else if (triggerValue is not null and not "timeline" &&
                         triggerValue != animation.GetProperty("start").GetString())
                    diagnostics.Add(new(
                        "ppj.animation.triggerMismatch",
                        "A non-timeline animation trigger must agree with the node start condition; use trigger=timeline to retain start as the source of truth.",
                        path + ".trigger"));
            }
        }
    }

    private static void ValidateTimingGraph(JsonElement page, string pagePath, List<PpjDiagnostic> diagnostics)
    {
        if (!page.TryGetProperty("timing", out var timing)) return;
        var nodes = timing.GetProperty("nodes");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < nodes.GetArrayLength(); index++)
        {
            var node = nodes[index];
            var nodePath = $"{pagePath}.timing.nodes[{index}]";
            var id = node.GetProperty("id").GetString()!;
            if (!ids.Add(id)) diagnostics.Add(new("ppj.id.duplicate", $"Timing node ID {id} is duplicated.", nodePath + ".id"));
            if (node.TryGetProperty("repeat", out var repeat) && repeat.GetInt32() is < 1 or > 8)
                diagnostics.Add(new("ppj.timing.repeat", "Timing repeat must be between 1 and 8.", nodePath + ".repeat"));
            if (node.TryGetProperty("easing", out var easing) && easing.GetString() is not (null or "linear" or "ease-in" or "ease-out" or "ease-in-out"))
                diagnostics.Add(new("ppj.timing.easing", "Timing easing is outside the bounded native profile.", nodePath + ".easing"));
            if (node.TryGetProperty("trigger", out var trigger))
            {
                var triggerValue = trigger.GetString();
                if (triggerValue is not (null or "timeline" or "onClick" or "afterPrevious" or "withPrevious"))
                    diagnostics.Add(new("ppj.timing.triggerUnsupported", "The timing trigger is outside the bounded click/previous profile.", nodePath + ".trigger"));
                else if (triggerValue is not null and not "timeline" &&
                         triggerValue != node.GetProperty("start").GetString())
                    diagnostics.Add(new(
                        "ppj.timing.triggerMismatch",
                        "A non-timeline timing trigger must agree with the node start condition; use trigger=timeline to retain start as the source of truth.",
                        nodePath + ".trigger"));
            }
        }
    }

    private static bool IsInlineStreamgraph(PpjChartElementModel chart) =>
        chart.ChartType == "area" &&
        chart.Raw.TryGetProperty("style", out var style) &&
        style.TryGetProperty("stacking", out var stacking) &&
        stacking.ValueKind == JsonValueKind.String &&
        stacking.GetString() == "stream";

    private static bool IsInlineSizedBubble(PpjChartElementModel chart) =>
        chart.ChartType == "bubble" &&
        chart.Raw.TryGetProperty("style", out var style) &&
        (style.TryGetProperty("bubbleSizeScale", out _) || style.TryGetProperty("bubbleRadiusRange", out _));

    private static bool IsInlinePictographicChart(PpjChartElementModel chart) =>
        chart.ChartType is "bar" or "column" &&
        chart.Data.Series.Any(series => series.Raw.TryGetProperty("symbol", out _));

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
            if (transition.Type == "none")
            {
                if (HasConfiguredTransitionFields(transition))
                    diagnostics.Add(new("ppj.transition.noneFields", "A none transition must not carry effect, timing, advance, or Morph fields.", path));
                continue;
            }

            if (PpjTransitionLowering.IsBaseEffect(transition.Type))
            {
                if (transition.FromPageId is not null || transition.MorphPairs.Count > 0)
                    diagnostics.Add(new("ppj.transition.morphFields", "fromPage and morphPairs are only valid for morph transitions.", path));
                if (!PpjTransitionLowering.TryBuildBase(transition, out _, out var error))
                    diagnostics.Add(new("ppj.transition.profile", error ?? $"Transition {transition.Type} is invalid.", path));
                continue;
            }

            if (transition.Type != "morph") continue;
            if (HasBaseOnlyTransitionFields(transition))
                diagnostics.Add(new("ppj.transition.baseFields", "direction, orientation, speed, throughBlack, spokes, and advance fields are not valid for Morph.", path));
            if (transition.MorphPairs.Count == 0)
                diagnostics.Add(new("ppj.transition.morphPairs", "Morph requires at least one object pair.", path + ".morphPairs"));

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

    private static bool HasConfiguredTransitionFields(PpjTransitionModel transition) =>
        transition.DurationMs is not null ||
        HasBaseOnlyTransitionFields(transition) ||
        transition.FromPageId is not null ||
        transition.MorphPairs.Count > 0;

    private static bool HasBaseOnlyTransitionFields(PpjTransitionModel transition) =>
        transition.Direction is not null ||
        transition.Orientation is not null ||
        transition.Speed is not null ||
        transition.ThroughBlack is not null ||
        transition.Spokes is not null ||
        transition.AdvanceOnClick is not null ||
        transition.AdvanceAfterMs is not null;

    private static void ValidatePresentationReferences(
        PpjProgramModel program,
        IReadOnlySet<string> pageIds,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, PpjElementModel>> pageElements,
        List<PpjDiagnostic> diagnostics)
    {
        for (var index = 0; index < program.Sections.Count; index++)
        {
            ValidatePageList(program.Sections[index].PageIds, pageIds, $"$.sections[{index}].pages", diagnostics);
            ValidateNativeRef(program.Sections[index].NativeRef, program.Source, $"$.sections[{index}].nativeRef", diagnostics);
        }
        for (var index = 0; index < program.CustomShows.Count; index++)
        {
            ValidatePageList(program.CustomShows[index].PageIds, pageIds, $"$.customShows[{index}].pages", diagnostics);
            ValidateNativeRef(program.CustomShows[index].NativeRef, program.Source, $"$.customShows[{index}].nativeRef", diagnostics);
        }

        var comments = program.Comments
            .GroupBy(item => item.Id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        for (var index = 0; index < program.Comments.Count; index++)
        {
            var comment = program.Comments[index];
            var path = $"$.comments[{index}]";
            if (comment.Kind is not ("legacy" or "modern"))
                diagnostics.Add(new("ppj.comment.kind", $"Comment kind {comment.Kind} is outside the bounded legacy/modern profile.", $"{path}.kind"));
            if (comment.Kind == "legacy" && comment.Status is not null)
                diagnostics.Add(new("ppj.comment.status", "Legacy comments cannot carry a modern status.", $"{path}.status"));
            if (comment.Kind == "modern")
            {
                if (comment.Status is not ("active" or "resolved" or "closed"))
                    diagnostics.Add(new("ppj.comment.status", "Modern comments require active, resolved, or closed status.", $"{path}.status"));
                else if (comment.Resolved != !comment.Status.Equals("active", StringComparison.Ordinal))
                    diagnostics.Add(new("ppj.comment.statusMismatch", "Modern comment resolved must agree with status (active=false; resolved/closed=true).", path));
                if (comment.ParentId is null)
                {
                    if (comment.TargetId is null)
                        diagnostics.Add(new("ppj.comment.target", "A modern root comment requires an element target.", $"{path}.target"));
                    if (!comment.Raw.TryGetProperty("anchor", out var anchor) || anchor.ValueKind != JsonValueKind.Object)
                    {
                        diagnostics.Add(new("ppj.comment.anchor", "A modern root comment requires an anchor.", $"{path}.anchor"));
                    }
                    else
                    {
                        var anchorKind = anchor.TryGetProperty("kind", out var kind) ? kind.GetString() : null;
                        var hasRange = anchor.TryGetProperty("textStart", out _) || anchor.TryGetProperty("textLength", out _);
                        if (anchorKind == "textRange" && (!anchor.TryGetProperty("textStart", out _) || !anchor.TryGetProperty("textLength", out _)))
                            diagnostics.Add(new("ppj.comment.anchor", "A textRange modern comment anchor requires textStart and textLength.", $"{path}.anchor"));
                        if (anchorKind == "element" && hasRange)
                            diagnostics.Add(new("ppj.comment.anchor", "An element modern comment anchor cannot carry text-range fields.", $"{path}.anchor"));
                    }
                }
                else
                {
                    if (comment.TargetId is not null)
                        diagnostics.Add(new("ppj.comment.target", "A modern reply inherits its root target and cannot declare target.", $"{path}.target"));
                    if (comment.Raw.TryGetProperty("anchor", out _))
                        diagnostics.Add(new("ppj.comment.anchor", "A modern reply inherits its root anchor and cannot declare anchor.", $"{path}.anchor"));
                }
            }
            if (!pageIds.Contains(comment.PageId))
                diagnostics.Add(new("ppj.comment.page", $"Comment page {comment.PageId} does not exist.", $"{path}.page"));
            else if (comment.TargetId is not null && !pageElements[comment.PageId].ContainsKey(comment.TargetId))
                diagnostics.Add(new("ppj.comment.target", $"Comment target {comment.TargetId} does not exist on page {comment.PageId}.", $"{path}.target"));
            if (comment.ParentId is not null && !comments.ContainsKey(comment.ParentId))
                diagnostics.Add(new("ppj.comment.parent", $"Parent comment {comment.ParentId} does not exist.", $"{path}.parent"));
            else if (comment.ParentId is not null && comments.TryGetValue(comment.ParentId, out var parent) &&
                     (parent.Kind != "modern" || parent.PageId != comment.PageId || parent.ParentId is not null))
                diagnostics.Add(new("ppj.comment.parent", "Modern replies must directly reference a root modern comment on the same page.", $"{path}.parent"));
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
        StringBuilder path,
        IReadOnlySet<string> assetIds,
        IReadOnlySet<string> colorIds,
        IReadOnlySet<string> fontIds,
        IReadOnlyDictionary<string, string> grammarTokenKinds,
        List<PpjDiagnostic> diagnostics)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            if (value.TryGetProperty("token", out var token) && token.ValueKind == JsonValueKind.String &&
                IsColorReferencePath(path) && !colorIds.Contains(token.GetString()!))
            {
                var tokenName = token.GetString()!;
                var code = grammarTokenKinds.TryGetValue(tokenName, out var kind)
                    ? "ppj.grammar.tokenKind"
                    : "ppj.colorRef";
                var message = kind is null
                    ? $"Color token {tokenName} does not exist."
                    : $"Color token {tokenName} declares kind {kind}, expected color.";
                diagnostics.Add(new(code, message, PathWithProperty(path, "token")));
            }
            if (value.TryGetProperty("font", out var font) && font.ValueKind == JsonValueKind.String && !fontIds.Contains(font.GetString()!))
                diagnostics.Add(new("ppj.fontRef", $"Font {font.GetString()} does not exist.", PathWithProperty(path, "font")));
            foreach (var property in value.EnumerateObject())
            {
                var length = AppendProperty(path, property.Name);
                if (property.Name is "asset" or "posterAsset" or "payloadAsset" or "previewAsset" &&
                    property.Value.ValueKind == JsonValueKind.String && !assetIds.Contains(property.Value.GetString()!))
                {
                    diagnostics.Add(new("ppj.assetRef", $"Asset {property.Value.GetString()} does not exist.", path.ToString()));
                }
                ValidateResourceReferences(property.Value, path, assetIds, colorIds, fontIds, grammarTokenKinds, diagnostics);
                path.Length = length;
            }
        }
        else if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                var length = AppendIndex(path, index++);
                ValidateResourceReferences(item, path, assetIds, colorIds, fontIds, grammarTokenKinds, diagnostics);
                path.Length = length;
            }
        }
    }

    private static bool IsColorReferencePath(StringBuilder path)
    {
        var value = path.ToString();
        return value.EndsWith(".color", StringComparison.Ordinal) ||
               value.EndsWith(".highlight", StringComparison.Ordinal) ||
               value.EndsWith(".missingFill", StringComparison.Ordinal);
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

    private static string PathWithProperty(StringBuilder path, string property)
    {
        var length = AppendProperty(path, property);
        var result = path.ToString();
        path.Length = length;
        return result;
    }

    private static string PathWithProperties(StringBuilder path, string first, string second)
    {
        var length = AppendProperty(path, first);
        AppendProperty(path, second);
        var result = path.ToString();
        path.Length = length;
        return result;
    }

    private static string PathWithGradientStopOffset(StringBuilder path, int stopIndex)
    {
        var length = AppendProperty(path, "gradient");
        AppendProperty(path, "stops");
        AppendIndex(path, stopIndex);
        AppendProperty(path, "offset");
        var result = path.ToString();
        path.Length = length;
        return result;
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
        "fill" or "stroke" or "opacity" => target is PpjShapeElementModel or PpjIconElementModel or PpjTextElementModel,
        "line.path" => target is PpjShapeElementModel { Type: "line" },
        "frame.x" or "frame.y" or "frame.width" or "frame.height" => true,
        "image.asset" or "image.crop" or "image.focus" => target is PpjImageElementModel,
        "chart.title" or "chart.data" or "chart.frame" => target is PpjChartElementModel,
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
