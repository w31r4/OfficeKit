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

    internal sealed record TimingRead(
        IReadOnlyList<PresentationAnimation> Animations,
        PresentationMorph? Morph,
        bool Present,
        bool Editable,
        bool Addable,
        string SemanticSha256);

    internal static TimingRead Read(P.Slide source, IReadOnlyDictionary<uint, string> elementIdsByNativeId)
    {
        var timing = source.Timing;
        var morph = ReadMorph(source, elementIdsByNativeId);
        var addable = timing is null && morph is null && source.ChildElements.All(child => child is P.CommonSlideData or P.ColorMapOverride or P.Transition);
        if (timing is null)
            return new([], morph, morph is not null, false, addable, SemanticHash([], morph));

        try
        {
            var root = XElement.Parse(timing.OuterXml, LoadOptions.PreserveWhitespace);
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
                var animation = new PresentationAnimation
                {
                    Id = $"anim-motion-{nativeId}", TargetId = targetId, TargetKind = "element", Effect = "fly",
                    Phase = "entrance", Start = start, Direction = direction, DurationMs = 500,
                };
                if (delay > 0) animation.DelayMs = delay;
                animations.Add(animation);
            }

            foreach (var behavior in root.Descendants(XName.Get("animScale", PNamespace)))
            {
                var target = behavior.Descendants(XName.Get("spTgt", PNamespace)).FirstOrDefault();
                if (!uint.TryParse(target?.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId)) return Opaque(root);
                var isPulse = behavior.Descendants(XName.Get("to", PNamespace)).Any(to => to.Attribute("x")?.Value == "110000");
                var (start, delay) = ParseStart(behavior);
                var duration = ParseDuration(behavior.Descendants(XName.Get("cTn", PNamespace)).FirstOrDefault()?.Attribute("dur")?.Value) ?? 500U;
                var animation = new PresentationAnimation
                {
                    Id = $"anim-scale-{nativeId}", TargetId = targetId, TargetKind = "element", Effect = isPulse ? "pulse" : "zoom",
                    Phase = isPulse ? "emphasis" : "entrance", Start = start, DurationMs = duration,
                };
                if (delay > 0) animation.DelayMs = delay;
                animations.Add(animation);
            }

            foreach (var build in root.Descendants(XName.Get("bldP", PNamespace)))
            {
                if (!uint.TryParse(build.Attribute("spid")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                    !elementIdsByNativeId.TryGetValue(nativeId, out var targetId)) return Opaque(root);
                UpsertBuildAnimation(animations, targetId, "element", null, build.Attribute("build")?.Value == "p" ? "paragraph" : "whole");
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
                    "seriesEl" => "seriesElement",
                    "categoryEl" => "categoryElement",
                    _ => "allAtOnce",
                };
                UpsertBuildAnimation(animations, targetId, "chart", value, null);
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
                    "seriesEl" => "seriesElement",
                    "categoryEl" => "categoryElement",
                    _ => "allAtOnce",
                };
                UpsertBuildAnimation(animations, targetId, "chart", value, null);
            }

            return new(animations, morph, true, true, false, SemanticHash(animations, morph));
        }
        catch
        {
            return Opaque(new XElement("timing"));
        }
    }

    private static TimingRead Opaque(XElement root) => new([], null, true, false, false, Hash(Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting))));

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

    private static void UpsertBuildAnimation(List<PresentationAnimation> animations, string targetId, string targetKind, string? chartBuild, string? textBuild)
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
        animations.Add(animation);
    }

    internal static bool CanAdd(P.Slide source) => source.Timing is null && source.ChildElements.All(child => child is P.CommonSlideData or P.ColorMapOverride or P.Transition);

    internal static string SemanticHash(IEnumerable<PresentationAnimation> animations, PresentationMorph? morph)
    {
        var semantic = string.Join("|", animations.Select(animation => string.Join(",", animation.Id, animation.TargetId, animation.Effect, animation.Phase, animation.Start, animation.Direction, animation.DurationMs, animation.ChartBuild, animation.TextBuild)));
        if (morph is not null) semantic += $"|morph:{morph.DurationMs}:{string.Join(";", morph.Pairs.Select(pair => string.Join(",", pair.Key, pair.FromId, pair.ToId)))}";
        return Hash(Encoding.UTF8.GetBytes(semantic));
    }

    internal static void Build(P.Slide target, PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        if (source.Animations.Count == 0 && source.Morph is null) return;
        Apply(target, source, nativeIdsByElementId, allowOpaqueReplacement: true);
    }

    internal static void Apply(P.Slide target, PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId, bool allowOpaqueReplacement)
    {
        Validate(source, nativeIdsByElementId);
        if (target.Timing is not null && !allowOpaqueReplacement)
            throw new CodecException("unsupported_presentation_timing_edit", "Imported presentation timing is opaque and cannot be replaced safely.");
        target.Timing?.Remove();
        if (source.Animations.Count == 0)
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
        // Morph is a transition extension (not a timing sidecar). Prefixing the
        // destination object names with !! is PowerPoint's stable by-object
        // pairing contract; the typed pairs remain in the task/artifact plan.
        var transition = target.Transition;
        if (transition is null)
        {
            transition = new P.Transition { Speed = P.TransitionSpeedValues.Medium, AdvanceOnClick = true };
            transition.Append(new P.FadeTransition());
            target.AddChild(transition, true);
        }
        const string p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        transition.AddNamespaceDeclaration("p14", p14Namespace);
        transition.SetAttribute(new OpenXmlAttribute("p14", "dur", p14Namespace, morph.DurationMs.ToString(CultureInfo.InvariantCulture)));
        var extension = new P.ExtensionListWithModification($"<p:extLst xmlns:p=\"{PNamespace}\" xmlns:p15=\"http://schemas.microsoft.com/office/powerpoint/2015/09/main\"><p:ext uri=\"{{officekit-morph-v1}}\"><p15:morph option=\"byObject\"/></p:ext></p:extLst>");
        transition.Append(extension);
        // The native IDs are validated here even though the p15 element only
        // carries the option; it prevents accepting a stale pair silently.
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
    }

    internal static void Validate(PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
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
        }
        if (source.Morph is not null)
        {
            if (source.Morph.DurationMs is 0 or > MaxDurationMilliseconds)
                throw new CodecException("invalid_presentation_morph", "Presentation Morph duration_ms must be from 1 through 60000.");
            foreach (var pair in source.Morph.Pairs)
                if (!nativeIdsByElementId.ContainsKey(pair.FromId) && !nativeIdsByElementId.ContainsKey(pair.ToId))
                    throw new CodecException("invalid_presentation_morph", $"Presentation Morph pair {pair.Key} does not resolve to a local object.");
        }
    }

    private static string BuildXml(PresentationSlide source, IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        var sb = new StringBuilder();
        // Build-list grpId values refer to timing groups (cTn@grpId), not
        // cTn@id. PowerPoint and the Open XML validator both accept the
        // canonical group zero used by authored chart/text build lists.
        sb.Append($"<p:timing xmlns:p=\"{PNamespace}\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:officekit=\"{OfficeKitNamespace}\"><p:tnLst><p:par><p:cTn id=\"1\" grpId=\"0\" dur=\"indefinite\" nodeType=\"tmRoot\"><p:childTnLst><p:seq concurrent=\"1\" nextAc=\"seek\"><p:cTn id=\"2\" grpId=\"0\" dur=\"indefinite\" nodeType=\"mainSeq\"><p:childTnLst>");
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
                    sb.Append($"<p:par><p:cTn id=\"{animationId}\" dur=\"{animation.DurationMs}\">{AnimationBehaviorXml(animation, nativeId, animationId)}</p:cTn></p:par>");
                sb.Append("</p:childTnLst></p:cTn></p:par>");
            }
            else
            {
                var (animation, nativeId, animationId) = group[0];
                var start = animation.Start == "onClick"
                    ? $"<p:stCondLst><p:cond evt=\"onClick\" delay=\"{animation.DelayMs}\"/></p:stCondLst>"
                    : $"<p:stCondLst><p:cond delay=\"{animation.DelayMs}\"/></p:stCondLst>";
                sb.Append($"<p:par><p:cTn id=\"{animationId}\" dur=\"{animation.DurationMs}\">{start}{AnimationBehaviorXml(animation, nativeId, animationId)}</p:cTn></p:par>");
            }
        }
        sb.Append("</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>");
        var builds = new StringBuilder();
        foreach (var animation in source.Animations.Where(animation => !string.IsNullOrEmpty(animation.TextBuild) || !string.IsNullOrEmpty(animation.ChartBuild)))
        {
            var nativeId = nativeIdsByElementId[animation.TargetId];
            if (!string.IsNullOrEmpty(animation.TextBuild))
                builds.Append($"<p:bldP spid=\"{nativeId}\" grpId=\"0\" build=\"{(animation.TextBuild == "paragraph" ? "p" : "whole")}\" bldLvl=\"1\"/>");
            if (!string.IsNullOrEmpty(animation.ChartBuild))
            {
                var build = animation.ChartBuild switch { "series" => "series", "category" => "category", "seriesElement" => "seriesEl", "categoryElement" => "categoryEl", _ => "allAtOnce" };
                builds.Append($"<p:bldGraphic spid=\"{nativeId}\" grpId=\"0\"><p:bldSub><a:bldChart bld=\"{build}\" animBg=\"1\"/></p:bldSub></p:bldGraphic>");
            }
        }
        if (builds.Length > 0) sb.Append($"<p:bldLst>{builds}</p:bldLst>");
        sb.Append("</p:timing>");
        return sb.ToString();
    }

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

    private static PresentationMorph? ReadMorph(P.Slide source, IReadOnlyDictionary<uint, string> elementIdsByNativeId)
    {
        try
        {
            const string p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
            var transition = source.Transition;
            var hasMorph = transition?.Descendants().Any(child => child.LocalName == "morph") == true;
            if (!hasMorph)
            {
                var root = XElement.Parse(source.OuterXml, LoadOptions.PreserveWhitespace);
                hasMorph = root.Descendants().Any(child => child.Name.LocalName == "morph");
                if (!hasMorph) return null;
            }
            var duration = ParseDuration(transition?.GetAttribute("dur", p14Namespace).Value) ?? 600U;
            var pairs = source.Descendants()
                .Where(element => element.LocalName == "cNvPr")
                .Select(element => (Name: element.GetAttribute("name", string.Empty).Value ?? string.Empty, Id: element.GetAttribute("id", string.Empty).Value ?? string.Empty))
                .Where(item => item.Name.StartsWith("!!", StringComparison.Ordinal) && uint.TryParse(item.Id, NumberStyles.None, CultureInfo.InvariantCulture, out _))
                .Select(item => new { Key = item.Name[2..], NativeId = uint.Parse(item.Id, CultureInfo.InvariantCulture) })
                .Where(item => elementIdsByNativeId.ContainsKey(item.NativeId))
                .Select(item => new PresentationMorphPair { Key = item.Key, FromId = elementIdsByNativeId[item.NativeId], ToId = elementIdsByNativeId[item.NativeId] })
                .ToList();
            return pairs.Count == 0 ? null : new PresentationMorph { DurationMs = duration, Pairs = { pairs } };
        }
        catch
        {
            return null;
        }
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

    private static string Escape(string value) => System.Security.SecurityElement.Escape(value) ?? string.Empty;
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
