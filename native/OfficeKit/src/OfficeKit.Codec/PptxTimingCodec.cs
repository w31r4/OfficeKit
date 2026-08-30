using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// The timing writer intentionally owns a small, valid PresentationML profile.
// Unknown timing graphs are observed as opaque and never replaced. This keeps
// the public animation API useful for authored decks without risking imported
// PowerPoint choreography that we cannot prove equivalent.
internal static class PptxTimingCodec
{
    private const string PNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string OfficeKitNamespace = "urn:officekit:motion";
    private const uint MaxDurationMilliseconds = 60_000;
    private const int MaxSemanticAnimations = 32;
    private const int MaxExpandedTimingNodes = 64;

    internal sealed record TimingRead(
        IReadOnlyList<PresentationAnimation> Animations,
        PresentationMorph? Morph,
        bool Present,
        bool Editable,
        bool Addable,
        string SemanticSha256);

    internal static TimingRead Read(P.Slide source, IReadOnlyDictionary<uint, string> elementIdsByNativeId) =>
        Read(source, elementIdsByNativeId, null, null, null);

    internal static TimingRead Read(
        P.Slide source,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        P.Slide? previousSource,
        IReadOnlyDictionary<uint, string>? previousElementIdsByNativeId,
        string? previousSlideId)
    {
        var timing = source.Timing;
        var morphPresent = HasMorph(source);
        var morph = ReadMorph(source, elementIdsByNativeId, previousSource, previousElementIdsByNativeId, previousSlideId);
        if (morphPresent && morph is null) return Opaque(source);
        var addable = timing is null && morph is null && source.ChildElements.All(child => child is P.CommonSlideData or P.ColorMapOverride or P.Transition);
        if (timing is null)
            return new([], morph, morph is not null, morph is not null, addable, SemanticHash([], morph));

        try
        {
            var root = XElement.Parse(timing.OuterXml, LoadOptions.PreserveWhitespace);
            // Authored media is restored from the embedded PPJ snapshot. A
            // third-party media timing graph remains source-owned because the
            // animation projection does not own its playback semantics.
            if (root.Descendants().Any(node => node.Name == XName.Get("audio", PNamespace) || node.Name == XName.Get("video", PNamespace)))
                return Opaque(root);
            var animations = new List<PresentationAnimation>();
            var index = 0;
            foreach (var behavior in root.Descendants(XName.Get("animEffect", PNamespace)))
            {
                var target = behavior.Descendants(XName.Get("spTgt", PNamespace)).FirstOrDefault();
                if (!uint.TryParse(target?.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId))
                    return Opaque(root);
                var filter = behavior.Attribute("filter")?.Value ?? "fade";
                var (effect, direction) = ParseFilter(filter);
                var phase = behavior.Attribute("transition")?.Value switch
                {
                    "out" => "exit",
                    "none" => "emphasis",
                    _ => "entrance",
                };
                var ctn = behavior.Descendants(XName.Get("cTn", PNamespace)).FirstOrDefault();
                var duration = ParseDuration(ctn?.Attribute("dur")?.Value) ?? 500U;
                var (start, delay) = ParseStart(behavior);
                var stagger = ParseStagger(behavior);
                var id = behavior.Ancestors(XName.Get("cTn", PNamespace)).Select(node => node.Attribute("id")?.Value).FirstOrDefault(value => !string.IsNullOrEmpty(value)) ?? $"{index + 1}";
                var animation = new PresentationAnimation
                {
                    Id = $"anim-{id}",
                    TargetId = targetId,
                    TargetKind = "element",
                    Effect = effect,
                    Phase = phase,
                    Start = start,
                    Direction = direction,
                    DurationMs = duration,
                };
                if (delay > 0) animation.DelayMs = delay;
                if (stagger > 0) animation.StaggerMs = stagger;
                animations.Add(animation);
                index++;
            }

            foreach (var behavior in root.Descendants(XName.Get("animMotion", PNamespace)))
            {
                var target = behavior.Descendants(XName.Get("spTgt", PNamespace)).FirstOrDefault();
                if (!uint.TryParse(target?.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId)) return Opaque(root);
                var direction = behavior.Attribute("officekit:direction")?.Value ?? "right";
                var (start, delay) = ParseStart(behavior);
                var stagger = ParseStagger(behavior);
                var animation = new PresentationAnimation
                {
                    Id = $"anim-motion-{nativeId}", TargetId = targetId, TargetKind = "element", Effect = "fly",
                    Phase = "entrance", Start = start, Direction = direction, DurationMs = 500,
                };
                if (delay > 0) animation.DelayMs = delay;
                if (stagger > 0) animation.StaggerMs = stagger;
                animations.Add(animation);
            }

            foreach (var behavior in root.Descendants(XName.Get("animScale", PNamespace)))
            {
                var target = behavior.Descendants(XName.Get("spTgt", PNamespace)).FirstOrDefault();
                if (!uint.TryParse(target?.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId)) return Opaque(root);
                var isPulse = behavior.Descendants(XName.Get("to", PNamespace)).Any(to => to.Attribute("x")?.Value == "110000");
                var (start, delay) = ParseStart(behavior);
                var stagger = ParseStagger(behavior);
                var duration = ParseDuration(behavior.Descendants(XName.Get("cTn", PNamespace)).FirstOrDefault()?.Attribute("dur")?.Value) ?? 500U;
                var animation = new PresentationAnimation
                {
                    Id = $"anim-scale-{nativeId}", TargetId = targetId, TargetKind = "element", Effect = isPulse ? "pulse" : "zoom",
                    Phase = isPulse ? "emphasis" : "entrance", Start = start, DurationMs = duration,
                };
                if (delay > 0) animation.DelayMs = delay;
                if (stagger > 0) animation.StaggerMs = stagger;
                animations.Add(animation);
            }

            foreach (var build in root.Descendants(XName.Get("bldP", PNamespace)))
            {
                if (!uint.TryParse(build.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId)) return Opaque(root);
                UpsertBuildAnimation(animations, targetId, "text", null, build.Attribute("build")?.Value == "p" ? "paragraph" : "whole", null);
            }

            var chartBuilds = root.Descendants(XName.Get("bldOleChart", PNamespace)).ToArray();
            foreach (var build in chartBuilds)
            {
                if (!uint.TryParse(build.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId)) return Opaque(root);
                var value = build.Attribute("bld")?.Value switch
                {
                    "series" => "series",
                    "category" => "category",
                    "seriesEl" => "series-element",
                    "categoryEl" => "category-element",
                    _ => "all-at-once",
                };
                UpsertBuildAnimation(animations, targetId, "chart", value, null, ParseBoolean(build.Attribute("animBg")?.Value));
            }
            foreach (var build in root.Descendants(XName.Get("bldGraphic", PNamespace)).Descendants(XName.Get("bldChart", "http://schemas.openxmlformats.org/drawingml/2006/main")))
            {
                var owner = build.Ancestors(XName.Get("bldGraphic", PNamespace)).FirstOrDefault();
                if (!uint.TryParse(owner?.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId)) return Opaque(root);
                var value = build.Attribute("bld")?.Value switch
                {
                    "series" => "series",
                    "category" => "category",
                    "seriesEl" => "series-element",
                    "categoryEl" => "category-element",
                    _ => "all-at-once",
                };
                UpsertBuildAnimation(animations, targetId, "chart", value, null, ParseBoolean(build.Attribute("animBg")?.Value));
            }

            return new(animations, morph, true, true, false, SemanticHash(animations, morph));
        }
        catch
        {
            return Opaque(source);
        }
    }

    private static TimingRead Opaque(XElement root) => new([], null, true, false, false, Hash(Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting))));

    private static TimingRead Opaque(P.Slide source)
    {
        var root = XElement.Parse(source.OuterXml, LoadOptions.PreserveWhitespace);
        var motionXml = string.Concat(root.Elements().Where(element =>
            element.Name.LocalName == "timing" || element.DescendantsAndSelf().Any(child => child.Name.LocalName == "morph"))
            .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        return new([], null, true, false, false, Hash(Encoding.UTF8.GetBytes(motionXml)));
    }

    private static (string Start, uint Delay) ParseStart(XElement behavior)
    {
        var group = behavior.Ancestors(XName.Get("cTn", PNamespace))
            .FirstOrDefault(node => node.Attribute("nodeType")?.Value == "withGroup")?
            .Ancestors(XName.Get("par", PNamespace)).FirstOrDefault();
        var groupBehaviors = group?.Descendants().Where(node => node.Name.LocalName is "animEffect" or "animScale" or "animMotion").ToList() ?? [];
        var withPrevious = groupBehaviors.IndexOf(behavior) > 0;
        var condition = behavior.Ancestors(XName.Get("cTn", PNamespace))
            .SelectMany(node => node.Elements(XName.Get("stCondLst", PNamespace)))
            .SelectMany(node => node.Elements(XName.Get("cond", PNamespace)))
            .FirstOrDefault();
        var delay = ParseDuration(condition?.Attribute("delay")?.Value) ?? 0U;
        if (withPrevious) return ("withPrevious", delay);
        return (string.Equals(condition?.Attribute("evt")?.Value, "onClick", StringComparison.Ordinal) ? "onClick" : "afterPrevious", delay);
    }

    private static uint ParseStagger(XElement behavior)
    {
        var iterate = behavior.Ancestors(XName.Get("cTn", PNamespace))
            .Select(node => node.Element(XName.Get("iterate", PNamespace)))
            .FirstOrDefault(node => node is not null);
        return ParseDuration(iterate?.Element(XName.Get("tmAbs", PNamespace))?.Attribute("val")?.Value) ?? 0U;
    }

    private static void UpsertBuildAnimation(List<PresentationAnimation> animations, string targetId, string targetKind, string? chartBuild, string? textBuild, bool? animateChartBackground)
    {
        // A build list is separate from the effect graph. When several effects
        // target the same object, bind the build to the first compatible effect
        // rather than whichever effect happened to be parsed last.
        var existing = chartBuild is not null
            ? animations.FirstOrDefault(animation => animation.TargetId == targetId && string.IsNullOrEmpty(animation.ChartBuild))
            : animations.FirstOrDefault(animation => animation.TargetId == targetId && string.IsNullOrEmpty(animation.TextBuild));
        if (existing is not null)
        {
            existing.TargetKind = targetKind;
            if (chartBuild is not null) existing.ChartBuild = chartBuild;
            if (textBuild is not null) existing.TextBuild = textBuild;
            if (animateChartBackground is not null) existing.AnimateChartBackground = animateChartBackground.Value;
            return;
        }
        var stableId = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(targetId))).ToLowerInvariant()[..8];
        var animation = new PresentationAnimation
        {
            Id = $"anim-build-{stableId}", TargetId = targetId, TargetKind = targetKind, Effect = "fade",
            Phase = "entrance", Start = "afterPrevious", DurationMs = 500,
        };
        if (chartBuild is not null) animation.ChartBuild = chartBuild;
        if (textBuild is not null) animation.TextBuild = textBuild;
        if (animateChartBackground is not null) animation.AnimateChartBackground = animateChartBackground.Value;
        animations.Add(animation);
    }

    internal static bool CanAdd(P.Slide source) => source.Timing is null && source.ChildElements.All(child => child is P.CommonSlideData or P.ColorMapOverride or P.Transition);

    internal static string SemanticHash(IEnumerable<PresentationAnimation> animations, PresentationMorph? morph)
    {
        var semantic = string.Join("|", animations.Select(animation => string.Join(",",
            animation.Id, animation.TargetId, animation.TargetKind, animation.Effect, animation.Phase,
            animation.Start, animation.Direction, animation.DurationMs, animation.DelayMs,
            animation.ChartBuild, animation.TextBuild, animation.StaggerMs,
            animation.HasAnimateChartBackground ? animation.AnimateChartBackground : null)));
        if (morph is not null) semantic += $"|morph:{morph.FromSlideId}:{morph.DurationMs}:{string.Join(";", morph.Pairs.Select(pair => string.Join(",", pair.Key, pair.FromId, pair.ToId)))}";
        return Hash(Encoding.UTF8.GetBytes(semantic));
    }

    internal static void Build(P.Slide target, PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        if (source.Animations.Count == 0 && !MediaElements(source.Elements).Any() && source.Morph is null) return;
        Apply(target, source, nativeIdsByElementId, allowOpaqueReplacement: true);
    }

    internal static void Apply(P.Slide target, PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId, bool allowOpaqueReplacement)
    {
        Validate(source, nativeIdsByElementId);
        if ((target.Timing is not null || HasMorph(target)) && !allowOpaqueReplacement)
            throw new CodecException("unsupported_presentation_timing_edit", "Imported presentation timing is opaque and cannot be replaced safely.");
        RemoveCanonicalMorph(target);
        target.Timing?.Remove();
        if (source.Animations.Count == 0 && !MediaElements(source.Elements).Any())
        {
            if (source.Morph is not null) ApplyMorphTransition(target, source.Morph, nativeIdsByElementId);
            return;
        }
        var xml = BuildXml(source, nativeIdsByElementId);
        target.AddChild(new P.Timing(xml), true);
        if (source.Morph is not null) ApplyMorphTransition(target, source.Morph, nativeIdsByElementId);
    }

    private static void ApplyMorphTransition(P.Slide target, PresentationMorph morph, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        // PowerPoint writes modern transitions through Markup Compatibility:
        // a p159 Morph choice plus a plain fade fallback. Stable !! names on
        // both adjacent slides provide the explicit by-object pairing.
        foreach (var pair in morph.Pairs)
        {
            if (!nativeIdsByElementId.ContainsKey(pair.ToId))
                throw new CodecException("invalid_presentation_morph", $"Morph pair {pair.Key} does not resolve to a slide object.");
            var nativeId = nativeIdsByElementId[pair.ToId];
            foreach (var drawingName in target.Descendants().Where(element => element.LocalName == "cNvPr"))
            {
                if (drawingName.GetAttribute("id", string.Empty).Value != nativeId.ToString(CultureInfo.InvariantCulture)) continue;
                drawingName.SetAttribute(new OpenXmlAttribute(string.Empty, "name", string.Empty, $"!!{pair.Key}"));
            }
        }
        var duration = morph.DurationMs.ToString(CultureInfo.InvariantCulture);
        var alternate = new AlternateContent($"<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice xmlns:p159=\"http://schemas.microsoft.com/office/powerpoint/2015/09/main\" Requires=\"p159\"><p:transition xmlns:p=\"{PNamespace}\" spd=\"slow\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" p14:dur=\"{duration}\"><p159:morph option=\"byObject\"/></p:transition></mc:Choice><mc:Fallback><p:transition xmlns:p=\"{PNamespace}\" spd=\"slow\"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>");
        var anchor = (OpenXmlElement?)target.ColorMapOverride ?? target.CommonSlideData;
        if (anchor is null) target.Append(alternate);
        else target.InsertAfter(alternate, anchor);
    }

    private static void RemoveCanonicalMorph(P.Slide target)
    {
        foreach (var alternate in target.Elements<AlternateContent>().Where(element =>
                     element.OuterXml.Contains("http://schemas.microsoft.com/office/powerpoint/2015/09/main", StringComparison.Ordinal) &&
                     element.OuterXml.Contains(":morph", StringComparison.Ordinal)).ToArray())
            alternate.Remove();
        foreach (var transition in target.Elements<P.Transition>().Where(element =>
                     element.Descendants().Any(child => child.LocalName == "morph" && child.NamespaceUri == "http://schemas.microsoft.com/office/powerpoint/2015/09/main")).ToArray())
            transition.Remove();
    }

    internal static void Validate(PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        if (source.Animations.Count > MaxSemanticAnimations || source.Animations.Count * 2 > MaxExpandedTimingNodes)
            throw new CodecException("presentation_animation_limit_exceeded", $"Presentation slides support at most {MaxSemanticAnimations} semantic animations and {MaxExpandedTimingNodes} expanded timing nodes.");
        foreach (var animation in source.Animations)
        {
            if (animation.DurationMs is 0 or > MaxDurationMilliseconds)
                throw new CodecException("invalid_presentation_animation", "Presentation animation duration_ms must be from 1 through 60000.");
            if (!nativeIdsByElementId.ContainsKey(animation.TargetId))
                throw new CodecException("invalid_presentation_animation_target", $"Presentation animation target {animation.TargetId} is not present on the slide.");
            if (animation.Effect is not ("fade" or "wipe" or "fly" or "zoom" or "pulse"))
                throw new CodecException("invalid_presentation_animation", $"Unsupported presentation animation effect {animation.Effect}.");
            if (animation.Start is not ("withPrevious" or "afterPrevious" or "onClick"))
                throw new CodecException("invalid_presentation_animation", "Presentation animation start must be withPrevious, afterPrevious, or onClick.");
            if (animation.DelayMs > MaxDurationMilliseconds || animation.StaggerMs > 10_000)
                throw new CodecException("invalid_presentation_animation", "Presentation animation delay or stagger exceeds the supported bound.");
            if (animation.Effect == "pulse" && animation.Phase != "emphasis" || animation.Effect != "pulse" && animation.Phase == "emphasis")
                throw new CodecException("invalid_presentation_animation", "Presentation emphasis supports pulse only, and pulse requires emphasis.");
            if (!string.IsNullOrEmpty(animation.TextBuild) && animation.TargetKind is not ("shape" or "textbox" or "text" or "element"))
                throw new CodecException("invalid_presentation_animation", "Presentation text builds require a text-bearing target.");
            if (!string.IsNullOrEmpty(animation.ChartBuild) && animation.TargetKind != "chart")
                throw new CodecException("invalid_presentation_animation", "Presentation chart builds require a chart target.");
            if (!string.IsNullOrEmpty(animation.ChartBuild) && animation.ChartBuild is not ("all-at-once" or "series" or "category" or "series-element" or "category-element"))
                throw new CodecException("invalid_presentation_animation", "Presentation chart build is unsupported.");
            if (animation.HasStaggerMs && animation.StaggerMs > 0 && animation.TextBuild != "paragraph" && animation.ChartBuild is not ("series" or "category" or "series-element" or "category-element"))
                throw new CodecException("invalid_presentation_animation", "Presentation stagger requires paragraph text or a segmented chart build.");
            if (animation.HasAnimateChartBackground && string.IsNullOrEmpty(animation.ChartBuild))
                throw new CodecException("invalid_presentation_animation", "Presentation chart-background animation requires a chart build.");
        }
        foreach (var element in MediaElements(source.Elements))
            if (!nativeIdsByElementId.ContainsKey(element.Id))
                throw new CodecException("invalid_presentation_media", $"Presentation media target {element.Id} is not present on the slide.");
        if (source.Morph is not null)
        {
            if (string.IsNullOrWhiteSpace(source.Morph.FromSlideId))
                throw new CodecException("invalid_presentation_morph", "Presentation Morph requires an adjacent from_slide_id.");
            if (source.Morph.DurationMs is 0 or > MaxDurationMilliseconds)
                throw new CodecException("invalid_presentation_morph", "Presentation Morph duration_ms must be from 1 through 60000.");
            foreach (var pair in source.Morph.Pairs)
                if (!nativeIdsByElementId.ContainsKey(pair.FromId) && !nativeIdsByElementId.ContainsKey(pair.ToId))
                    throw new CodecException("invalid_presentation_morph", $"Presentation Morph pair {pair.Key} does not resolve to a local object.");
        }
    }

    internal static void ValidateMorphContext(PresentationSlide destination, PresentationSlide? source)
    {
        var morph = destination.Morph;
        if (morph is null) return;
        if (source is null || !morph.FromSlideId.Equals(source.Id, StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_morph", "Presentation Morph must reference the immediately preceding slide.");
        if (morph.Pairs.Count is < 1 or > 256)
            throw new CodecException("invalid_presentation_morph", "Presentation Morph requires one through 256 object pairs.");

        var sourceElements = MorphElementIndex(source);
        var destinationElements = MorphElementIndex(destination);
        var sourceIds = new HashSet<string>(StringComparer.Ordinal);
        var destinationIds = new HashSet<string>(StringComparer.Ordinal);
        var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in morph.Pairs)
        {
            if (string.IsNullOrWhiteSpace(pair.Key) || !keys.Add(pair.Key) ||
                !sourceIds.Add(pair.FromId) || !destinationIds.Add(pair.ToId))
                throw new CodecException("invalid_presentation_morph", "Presentation Morph pair keys and object targets must be unique.");
            if (!sourceElements.TryGetValue(pair.FromId, out var from) || !destinationElements.TryGetValue(pair.ToId, out var to))
                throw new CodecException("invalid_presentation_morph", $"Presentation Morph pair {pair.Key} does not resolve across the adjacent slides.");
            if (!MorphCompatible(from, to))
                throw new CodecException("invalid_presentation_morph", $"Presentation Morph pair {pair.Key} requires compatible non-chart objects.");
            var expectedName = $"!!{pair.Key}";
            if (!from.Name.Equals(expectedName, StringComparison.Ordinal) || !to.Name.Equals(expectedName, StringComparison.Ordinal))
                throw new CodecException("invalid_presentation_morph", $"Presentation Morph pair {pair.Key} must use matching Selection Pane names on both slides.");
        }
    }

    private static IReadOnlyDictionary<string, PresentationElement> MorphElementIndex(PresentationSlide slide)
    {
        var output = new Dictionary<string, PresentationElement>(StringComparer.Ordinal);
        foreach (var element in FlattenElements(slide.Elements))
            if (string.IsNullOrWhiteSpace(element.Id) || !output.TryAdd(element.Id, element))
                throw new CodecException("invalid_presentation_morph", "Presentation Morph slide objects require unique non-empty identities.");
        return output;
    }

    private static IEnumerable<PresentationElement> FlattenElements(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.ContentCase != PresentationElement.ContentOneofCase.Group) continue;
            foreach (var child in FlattenElements(element.Group.Children)) yield return child;
        }
    }

    private static bool MorphCompatible(PresentationElement from, PresentationElement to) =>
        from.ContentCase == to.ContentCase &&
        from.ContentCase is PresentationElement.ContentOneofCase.Shape or
            PresentationElement.ContentOneofCase.Image or
            PresentationElement.ContentOneofCase.Table or
            PresentationElement.ContentOneofCase.Connector or
            PresentationElement.ContentOneofCase.Group;

    private static string BuildXml(PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        var sb = new StringBuilder();
        // Build-list grpId values refer to timing groups (cTn@grpId), not
        // cTn@id. PowerPoint and the Open XML validator both accept the
        // canonical group zero used by authored chart/text build lists.
        sb.Append($"<p:timing xmlns:p=\"{PNamespace}\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:officekit=\"{OfficeKitNamespace}\"><p:tnLst><p:par><p:cTn id=\"1\" grpId=\"0\" dur=\"indefinite\" nodeType=\"tmRoot\"><p:childTnLst>");
        var mediaNodeId = 100U;
        foreach (var element in MediaElements(source.Elements))
        {
            var media = element.Media;
            var mute = media.HasMute && media.Mute ? " mute=\"1\"" : string.Empty;
            var repeat = media.HasLoop && media.Loop ? " repeatCount=\"indefinite\"" : string.Empty;
            var kind = media.MediaType == "audio" ? "audio" : "video";
            sb.Append($"<p:{kind}><p:cMediaNode vol=\"80000\"{mute}><p:cTn id=\"{mediaNodeId++}\" fill=\"hold\" display=\"0\"{repeat}><p:stCondLst><p:cond delay=\"indefinite\"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid=\"{nativeIdsByElementId[element.Id]}\"/></p:tgtEl></p:cMediaNode></p:{kind}>");
        }
        if (source.Animations.Count > 0)
            sb.Append("<p:seq concurrent=\"1\" nextAc=\"seek\"><p:cTn id=\"2\" grpId=\"0\" dur=\"indefinite\" nodeType=\"mainSeq\"><p:childTnLst>");
        var id = 10U;
        var groupId = 3U;
        var groups = new List<List<(PresentationAnimation Animation, uint NativeId, uint Id)>>();
        foreach (var animation in source.Animations)
        {
            var item = (animation, nativeIdsByElementId[animation.TargetId], id);
            if (animation.Start == "withPrevious" && groups.Count > 0)
                groups[^1].Add(item);
            else
                groups.Add([item]);
            id += 2;
        }
        foreach (var group in groups)
        {
            if (group.Count > 1)
            {
                var first = group[0].Animation;
                var groupStart = first.Start == "onClick"
                    ? $"<p:stCondLst><p:cond evt=\"onClick\" delay=\"{first.DelayMs}\"/></p:stCondLst>"
                    : $"<p:stCondLst><p:cond delay=\"{first.DelayMs}\"/></p:stCondLst>";
                sb.Append($"<p:par><p:cTn id=\"{groupId++}\" dur=\"indefinite\" nodeType=\"withGroup\">{groupStart}<p:childTnLst>");
                foreach (var (animation, nativeId, animationId) in group)
                    sb.Append($"<p:par><p:cTn id=\"{animationId}\" dur=\"{animation.DurationMs}\">{IterateXml(animation)}{AnimationBehaviorXml(animation, nativeId, animationId)}</p:cTn></p:par>");
                sb.Append("</p:childTnLst></p:cTn></p:par>");
            }
            else
            {
                var (animation, nativeId, animationId) = group[0];
                var start = animation.Start == "onClick"
                    ? $"<p:stCondLst><p:cond evt=\"onClick\" delay=\"{animation.DelayMs}\"/></p:stCondLst>"
                    : $"<p:stCondLst><p:cond delay=\"{animation.DelayMs}\"/></p:stCondLst>";
                sb.Append($"<p:par><p:cTn id=\"{animationId}\" dur=\"{animation.DurationMs}\">{start}{IterateXml(animation)}{AnimationBehaviorXml(animation, nativeId, animationId)}</p:cTn></p:par>");
            }
        }
        if (source.Animations.Count > 0)
            sb.Append("</p:childTnLst></p:cTn></p:seq>");
        sb.Append("</p:childTnLst></p:cTn></p:par></p:tnLst>");
        var builds = new StringBuilder();
        foreach (var animation in source.Animations.Where(animation => !string.IsNullOrEmpty(animation.TextBuild) || !string.IsNullOrEmpty(animation.ChartBuild)))
        {
            var nativeId = nativeIdsByElementId[animation.TargetId];
            if (!string.IsNullOrEmpty(animation.TextBuild))
                builds.Append($"<p:bldP spid=\"{nativeId}\" grpId=\"0\" build=\"{(animation.TextBuild == "paragraph" ? "p" : "whole")}\" bldLvl=\"1\"/>");
            if (!string.IsNullOrEmpty(animation.ChartBuild))
            {
                var build = animation.ChartBuild switch { "series" => "series", "category" => "category", "series-element" => "seriesEl", "category-element" => "categoryEl", _ => "allAtOnce" };
                var animateBackground = animation.HasAnimateChartBackground && animation.AnimateChartBackground ? "1" : "0";
                builds.Append($"<p:bldGraphic spid=\"{nativeId}\" grpId=\"0\"><p:bldSub><a:bldChart bld=\"{build}\" animBg=\"{animateBackground}\"/></p:bldSub></p:bldGraphic>");
            }
        }
        if (builds.Length > 0) sb.Append($"<p:bldLst>{builds}</p:bldLst>");
        sb.Append("</p:timing>");
        return sb.ToString();
    }

    private static IEnumerable<PresentationElement> MediaElements(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            if (element.ContentCase == PresentationElement.ContentOneofCase.Media) yield return element;
            if (element.ContentCase != PresentationElement.ContentOneofCase.Group) continue;
            foreach (var child in MediaElements(element.Group.Children)) yield return child;
        }
    }

    private static string IterateXml(PresentationAnimation animation) =>
        animation.HasStaggerMs && animation.StaggerMs > 0
            ? $"<p:iterate type=\"el\"><p:tmAbs val=\"{animation.StaggerMs}\"/></p:iterate>"
            : string.Empty;

    private static string AnimationBehaviorXml(PresentationAnimation animation, uint nativeId, uint animationId)
    {
        var phase = animation.Phase == "exit" ? "out" : animation.Phase == "emphasis" ? "none" : "in";
        var filter = animation.Effect switch
        {
            "wipe" => $"wipe({animation.Direction})",
            "fly" => $"fly({animation.Direction})",
            "zoom" => "zoom",
            "pulse" => "pulse",
            _ => "fade",
        };
        var behavior = animation.Effect is "zoom" or "pulse"
            ? $"<p:animScale zoomContents=\"1\"><p:cBhvr><p:cTn id=\"{animationId + 1}\" dur=\"{animation.DurationMs}\"/><p:tgtEl><p:spTgt spid=\"{nativeId}\"/></p:tgtEl></p:cBhvr><p:from x=\"{(animation.Effect == "pulse" ? "100000" : "0")}\" y=\"{(animation.Effect == "pulse" ? "100000" : "0")}\"/><p:to x=\"{(animation.Effect == "pulse" ? "110000" : "100000")}\" y=\"{(animation.Effect == "pulse" ? "110000" : "100000")}\"/></p:animScale>"
            : $"<p:animEffect transition=\"{phase}\" filter=\"{Escape(filter)}\"><p:cBhvr><p:cTn id=\"{animationId + 1}\" dur=\"{animation.DurationMs}\"/><p:tgtEl><p:spTgt spid=\"{nativeId}\"/></p:tgtEl></p:cBhvr></p:animEffect>";
        return $"<p:childTnLst>{behavior}</p:childTnLst>";
    }

    internal static bool HasMorph(P.Slide source)
    {
        try
        {
            var root = XElement.Parse(source.OuterXml, LoadOptions.PreserveWhitespace);
            return root.Descendants().Any(child =>
                child.Name.LocalName == "morph" && child.Name.NamespaceName == "http://schemas.microsoft.com/office/powerpoint/2015/09/main");
        }
        catch
        {
            return false;
        }
    }

    private static PresentationMorph? ReadMorph(
        P.Slide source,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        P.Slide? previousSource,
        IReadOnlyDictionary<uint, string>? previousElementIdsByNativeId,
        string? previousSlideId)
    {
        try
        {
            const string p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
            if (!HasMorph(source)) return null;
            if (previousSource is null || previousElementIdsByNativeId is null || string.IsNullOrWhiteSpace(previousSlideId)) return null;
            var root = XElement.Parse(source.OuterXml, LoadOptions.PreserveWhitespace);
            var morph = root.Descendants().FirstOrDefault(child =>
                child.Name.LocalName == "morph" && child.Name.NamespaceName == "http://schemas.microsoft.com/office/powerpoint/2015/09/main");
            if (morph?.Attribute("option")?.Value != "byObject") return null;
            var transition = morph.Ancestors().FirstOrDefault(element => element.Name.LocalName == "transition");
            var duration = ParseDuration(transition?.Attribute(XName.Get("dur", p14Namespace))?.Value) ?? 600U;
            var destinations = MorphNames(source, elementIdsByNativeId);
            var sources = MorphNames(previousSource, previousElementIdsByNativeId);
            if (destinations is null || sources is null || destinations.Count == 0 || destinations.Keys.Any(key => !sources.ContainsKey(key))) return null;
            var pairs = destinations
                .OrderBy(item => item.Key, StringComparer.Ordinal)
                .Select(item => new PresentationMorphPair { Key = item.Key, FromId = sources[item.Key], ToId = item.Value })
                .ToList();
            return new PresentationMorph { FromSlideId = previousSlideId, DurationMs = duration, Pairs = { pairs } };
        }
        catch
        {
            return null;
        }
    }

    private static Dictionary<string, string>? MorphNames(P.Slide source, IReadOnlyDictionary<uint, string> elementIdsByNativeId)
    {
        var pairs = source.Descendants()
                .Where(element => element.LocalName == "cNvPr")
                .Select(element => (Name: element.GetAttribute("name", string.Empty).Value ?? string.Empty, Id: element.GetAttribute("id", string.Empty).Value ?? string.Empty))
                .Where(item => item.Name.StartsWith("!!", StringComparison.Ordinal) && uint.TryParse(item.Id, NumberStyles.None, CultureInfo.InvariantCulture, out _))
                .Select(item => new { Key = item.Name[2..], NativeId = uint.Parse(item.Id, CultureInfo.InvariantCulture) })
                .Where(item => elementIdsByNativeId.ContainsKey(item.NativeId))
                .ToList();
        if (pairs.Any(item => item.Key.Length == 0) || pairs.Select(item => item.Key).Distinct(StringComparer.Ordinal).Count() != pairs.Count)
        {
            return null;
        }
        return pairs.ToDictionary(item => item.Key, item => elementIdsByNativeId[item.NativeId], StringComparer.Ordinal);
    }

    private static (string Effect, string Direction) ParseFilter(string filter)
    {
        var open = filter.IndexOf('(');
        if (open < 0) return (filter switch { "zoom" => "zoom", "pulse" => "pulse", _ => "fade" }, "");
        var effect = filter[..open];
        var direction = filter[(open + 1)..].TrimEnd(')');
        return effect switch { "wipe" => ("wipe", direction), "fly" => ("fly", direction), _ => ("fade", "") };
    }

    private static uint? ParseDuration(string? value) => uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var number) ? number : null;

    private static bool? ParseBoolean(string? value) => value switch
    {
        "1" or "true" => true,
        "0" or "false" => false,
        _ => null,
    };

    private static string Escape(string value) => System.Security.SecurityElement.Escape(value) ?? string.Empty;
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
