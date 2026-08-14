using System.Security.Cryptography;
using System.Text;
using System.Globalization;
using DocumentFormat.OpenXml;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns one direct p:transition leaf on one SlidePart. This covers the complete
// ECMA-376 base-namespace effect vocabulary, but deliberately does not model
// p:timing, sound actions, or Office-version extension effects. The one
// Office 2010+ p14:dur leaf is projected as canonical integer milliseconds.
internal static class PptxTransitionCodec
{
    private const string PowerPoint2010Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
    internal const uint MaxDurationMilliseconds = 86_400_000;
    internal const uint MaxAdvanceAfterMilliseconds = 86_400_000;

    private enum TransitionEffectShape
    {
        Empty,
        OptionalBlack,
        Orientation,
        SlideDirection,
        EightDirection,
        CornerDirection,
        Split,
        Wheel,
        InOut,
    }

    private sealed record TransitionEffectProfile(
        string Name,
        Type ElementType,
        TransitionEffectShape Shape,
        Func<OpenXmlElement> Create);

    private static readonly TransitionEffectProfile[] EffectProfiles =
    [
        new("blinds", typeof(P.BlindsTransition), TransitionEffectShape.Orientation, static () => new P.BlindsTransition()),
        new("checker", typeof(P.CheckerTransition), TransitionEffectShape.Orientation, static () => new P.CheckerTransition()),
        new("circle", typeof(P.CircleTransition), TransitionEffectShape.Empty, static () => new P.CircleTransition()),
        new("comb", typeof(P.CombTransition), TransitionEffectShape.Orientation, static () => new P.CombTransition()),
        new("cover", typeof(P.CoverTransition), TransitionEffectShape.EightDirection, static () => new P.CoverTransition()),
        new("cut", typeof(P.CutTransition), TransitionEffectShape.OptionalBlack, static () => new P.CutTransition()),
        new("diamond", typeof(P.DiamondTransition), TransitionEffectShape.Empty, static () => new P.DiamondTransition()),
        new("dissolve", typeof(P.DissolveTransition), TransitionEffectShape.Empty, static () => new P.DissolveTransition()),
        new("fade", typeof(P.FadeTransition), TransitionEffectShape.OptionalBlack, static () => new P.FadeTransition()),
        new("newsflash", typeof(P.NewsflashTransition), TransitionEffectShape.Empty, static () => new P.NewsflashTransition()),
        new("plus", typeof(P.PlusTransition), TransitionEffectShape.Empty, static () => new P.PlusTransition()),
        new("pull", typeof(P.PullTransition), TransitionEffectShape.EightDirection, static () => new P.PullTransition()),
        new("push", typeof(P.PushTransition), TransitionEffectShape.SlideDirection, static () => new P.PushTransition()),
        new("random", typeof(P.RandomTransition), TransitionEffectShape.Empty, static () => new P.RandomTransition()),
        new("randomBar", typeof(P.RandomBarTransition), TransitionEffectShape.Orientation, static () => new P.RandomBarTransition()),
        new("split", typeof(P.SplitTransition), TransitionEffectShape.Split, static () => new P.SplitTransition()),
        new("strips", typeof(P.StripsTransition), TransitionEffectShape.CornerDirection, static () => new P.StripsTransition()),
        new("wedge", typeof(P.WedgeTransition), TransitionEffectShape.Empty, static () => new P.WedgeTransition()),
        new("wheel", typeof(P.WheelTransition), TransitionEffectShape.Wheel, static () => new P.WheelTransition()),
        new("wipe", typeof(P.WipeTransition), TransitionEffectShape.SlideDirection, static () => new P.WipeTransition()),
        new("zoom", typeof(P.ZoomTransition), TransitionEffectShape.InOut, static () => new P.ZoomTransition()),
    ];

    private static readonly IReadOnlyDictionary<string, TransitionEffectProfile> EffectProfilesByName =
        EffectProfiles.ToDictionary(profile => profile.Name, StringComparer.Ordinal);

    internal static PresentationTransition? Read(P.Slide source)
    {
        var transitions = source.Elements<P.Transition>().ToArray();
        return transitions.Length == 1 && TryRead(transitions[0], out var semantic) ? semantic : null;
    }

    internal static bool HasTransition(P.Slide source) => source.Elements<P.Transition>().Any();

    internal static bool Supports(P.Slide source)
    {
        var transitions = source.Elements<P.Transition>().ToArray();
        return transitions.Length == 1 &&
            TryRead(transitions[0], out _) &&
            source.ChildElements.All(child => child is P.CommonSlideData or P.ColorMapOverride or P.Transition);
    }

    // Adding a direct transition is narrower than merely observing that one
    // is absent. Any timing tree or extension can alter slideshow semantics,
    // so a source-bound add is allowed only on the ordinary Slide root shape.
    internal static bool CanAdd(P.Slide source) =>
        !HasTransition(source) &&
        source.ChildElements.All(child => child is P.CommonSlideData or P.ColorMapOverride);

    internal static void Validate(PresentationTransition? source)
    {
        if (source is null) return;
        if (!source.HasAdvanceOnClick)
            throw Invalid("Presentation transition requires an explicit advance_on_click value.");
        if (source.HasDurationMs && source.DurationMs > MaxDurationMilliseconds)
            throw Invalid($"Presentation transition duration_ms must not exceed {MaxDurationMilliseconds}.");
        if (source.HasAdvanceAfterMs && source.AdvanceAfterMs > MaxAdvanceAfterMilliseconds)
            throw Invalid($"Presentation transition advance_after_ms must not exceed {MaxAdvanceAfterMilliseconds}.");
        if (!IsSpeed(source.Speed))
            throw Invalid("Presentation transition speed must be slow, medium, or fast.");
        if (!EffectProfilesByName.TryGetValue(source.Effect, out var profile))
            throw Invalid($"Presentation transition effect must be one of: {string.Join(", ", EffectProfiles.Select(candidate => candidate.Name))}.");

        var directionRequired = profile.Shape is TransitionEffectShape.SlideDirection or TransitionEffectShape.EightDirection or
            TransitionEffectShape.CornerDirection or TransitionEffectShape.Split or TransitionEffectShape.InOut;
        if (directionRequired)
        {
            if (!IsDirection(profile.Shape, source.Direction))
                throw Invalid($"Presentation {profile.Name} transition direction must be {DirectionLabel(profile.Shape)}.");
        }
        else if (!string.IsNullOrEmpty(source.Direction))
            throw Invalid($"Presentation {profile.Name} transition must not carry direction.");

        var orientationRequired = profile.Shape is TransitionEffectShape.Orientation or TransitionEffectShape.Split;
        if (orientationRequired)
        {
            if (!IsOrientation(source.Orientation))
                throw Invalid($"Presentation {profile.Name} transition orientation must be horizontal or vertical.");
        }
        else if (!string.IsNullOrEmpty(source.Orientation))
            throw Invalid($"Presentation {profile.Name} transition must not carry orientation.");

        if (profile.Shape != TransitionEffectShape.OptionalBlack && source.HasThroughBlack)
            throw Invalid($"Presentation {profile.Name} transition must not carry through_black.");
        if (profile.Shape == TransitionEffectShape.Wheel)
        {
            if (!source.HasSpokes || source.Spokes is < 1 or > 8)
                throw Invalid("Presentation wheel transition spokes must be from 1 through 8.");
        }
        else if (source.HasSpokes)
            throw Invalid($"Presentation {profile.Name} transition must not carry spokes.");
    }

    internal static void Build(P.Slide target, PresentationTransition? source)
    {
        if (source is null) return;
        target.AddChild(BuildElement(source), true);
    }

    internal static void Apply(P.Slide target, PresentationTransition? source)
    {
        Validate(source);
        var transitions = target.Elements<P.Transition>().ToArray();
        if (transitions.Length > 1)
            throw new CodecException("presentation_transition_topology_changed", "Slide contains multiple transition elements.");
        var current = transitions.SingleOrDefault();
        if (source is null)
        {
            current?.Remove();
            return;
        }
        var replacement = BuildElement(source);
        if (current is null)
        {
            target.AddChild(replacement, true);
            return;
        }
        current.InsertAfterSelf(replacement);
        current.Remove();
    }

    internal static string SemanticHash(PresentationTransition? source) =>
        Hash((source?.Clone() ?? new PresentationTransition()).ToByteArray());

    // Used only for a no-op proof. Semantic hashing intentionally maps opaque
    // transitions to absence; this raw hash makes sure that an unsupported
    // native transition was not silently dropped or rewritten.
    internal static string ElementHash(P.Slide source)
    {
        var xml = string.Concat(source.Elements<P.Transition>().Select(transition => transition.OuterXml));
        return Hash(Encoding.UTF8.GetBytes(xml));
    }

    private static bool TryRead(P.Transition source, out PresentationTransition semantic)
    {
        semantic = new PresentationTransition();
        if (source.ExtendedAttributes.Any() || !HasTransitionAttributes(source) ||
            source.Speed?.Value is not { } speed || source.AdvanceOnClick?.Value is not { } advanceOnClick ||
            !TrySpeed(speed, out var speedName) || source.ChildElements.Count != 1)
            return false;
        semantic.Speed = speedName;
        semantic.AdvanceOnClick = advanceOnClick;
        if (source.Duration?.Value is { } durationText)
        {
            if (!TryMilliseconds(durationText, MaxDurationMilliseconds, out var duration)) return false;
            semantic.DurationMs = duration;
        }
        if (source.AdvanceAfterTime?.Value is { } advanceAfterText)
        {
            if (!TryMilliseconds(advanceAfterText, MaxAdvanceAfterMilliseconds, out var advanceAfter)) return false;
            semantic.AdvanceAfterMs = advanceAfter;
        }
        var child = source.FirstChild;
        var profile = EffectProfiles.FirstOrDefault(candidate => child?.GetType() == candidate.ElementType);
        if (profile is null || child is null) return false;
        semantic.Effect = profile.Name;
        return TryReadEffect(profile, child, semantic);
    }

    private static P.Transition BuildElement(PresentationTransition source)
    {
        Validate(source);
        var transition = new P.Transition
        {
            Speed = Speed(source.Speed),
            AdvanceOnClick = source.AdvanceOnClick,
        };
        if (source.HasDurationMs) transition.Duration = source.DurationMs.ToString(CultureInfo.InvariantCulture);
        if (source.HasAdvanceAfterMs) transition.AdvanceAfterTime = source.AdvanceAfterMs.ToString(CultureInfo.InvariantCulture);
        transition.Append(BuildEffect(EffectProfilesByName[source.Effect], source));
        return transition;
    }

    private static bool TryReadEffect(TransitionEffectProfile profile, OpenXmlElement source, PresentationTransition semantic)
    {
        switch (profile.Shape)
        {
            case TransitionEffectShape.Empty:
                return IsEmpty(source);
            case TransitionEffectShape.OptionalBlack:
                return TryOptionalBlack(source, semantic);
            case TransitionEffectShape.Orientation:
                if (!TrySingleAttribute(source, "dir", out var orientationToken) || !TryOrientationToken(orientationToken, out var orientation)) return false;
                semantic.Orientation = orientation;
                return true;
            case TransitionEffectShape.SlideDirection:
            case TransitionEffectShape.EightDirection:
            case TransitionEffectShape.CornerDirection:
            case TransitionEffectShape.InOut:
                if (!TrySingleAttribute(source, "dir", out var directionToken) || !TryDirectionToken(profile.Shape, directionToken, out var direction)) return false;
                semantic.Direction = direction;
                return true;
            case TransitionEffectShape.Split:
                if (!TryTwoAttributes(source, "orient", "dir", out var splitAttributes) ||
                    !TryOrientationToken(splitAttributes["orient"], out var splitOrientation) ||
                    !TryDirectionToken(profile.Shape, splitAttributes["dir"], out var splitDirection)) return false;
                semantic.Orientation = splitOrientation;
                semantic.Direction = splitDirection;
                return true;
            case TransitionEffectShape.Wheel:
                if (!TrySingleAttribute(source, "spokes", out var spokesToken) ||
                    !uint.TryParse(spokesToken, NumberStyles.None, CultureInfo.InvariantCulture, out var spokes) || spokes is < 1 or > 8) return false;
                semantic.Spokes = spokes;
                return true;
            default:
                return false;
        }
    }

    private static OpenXmlElement BuildEffect(TransitionEffectProfile profile, PresentationTransition source)
    {
        var element = profile.Create();
        switch (profile.Shape)
        {
            case TransitionEffectShape.OptionalBlack when source.HasThroughBlack:
                SetAttribute(element, "thruBlk", source.ThroughBlack ? "1" : "0");
                break;
            case TransitionEffectShape.Orientation:
                SetAttribute(element, "dir", OrientationToken(source.Orientation));
                break;
            case TransitionEffectShape.SlideDirection:
            case TransitionEffectShape.EightDirection:
            case TransitionEffectShape.CornerDirection:
            case TransitionEffectShape.InOut:
                SetAttribute(element, "dir", DirectionToken(profile.Shape, source.Direction));
                break;
            case TransitionEffectShape.Split:
                SetAttribute(element, "orient", OrientationToken(source.Orientation));
                SetAttribute(element, "dir", DirectionToken(profile.Shape, source.Direction));
                break;
            case TransitionEffectShape.Wheel:
                SetAttribute(element, "spokes", source.Spokes.ToString(CultureInfo.InvariantCulture));
                break;
        }
        return element;
    }

    private static bool IsEmpty(OpenXmlElement source) =>
        !source.ExtendedAttributes.Any() && source.GetAttributes().Count == 0 && source.ChildElements.Count == 0;

    private static bool TryOptionalBlack(OpenXmlElement source, PresentationTransition semantic)
    {
        if (source.ExtendedAttributes.Any() || source.ChildElements.Count != 0) return false;
        var attributes = source.GetAttributes();
        if (attributes.Count == 0) return true;
        if (attributes.Count != 1 || attributes[0].NamespaceUri.Length != 0 || attributes[0].LocalName != "thruBlk" ||
            !TryBooleanToken(attributes[0].Value ?? string.Empty, out var throughBlack)) return false;
        semantic.ThroughBlack = throughBlack;
        return true;
    }

    private static bool TrySingleAttribute(OpenXmlElement source, string name, out string value)
    {
        value = string.Empty;
        if (source.ExtendedAttributes.Any() || source.ChildElements.Count != 0 || !HasOnlyAttributes(source, name)) return false;
        value = source.GetAttribute(name, string.Empty).Value ?? string.Empty;
        return value.Length > 0;
    }

    private static bool TryTwoAttributes(OpenXmlElement source, string first, string second, out IReadOnlyDictionary<string, string> values)
    {
        values = new Dictionary<string, string>(StringComparer.Ordinal);
        if (source.ExtendedAttributes.Any() || source.ChildElements.Count != 0 || !HasOnlyAttributes(source, first, second)) return false;
        var output = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [first] = source.GetAttribute(first, string.Empty).Value ?? string.Empty,
            [second] = source.GetAttribute(second, string.Empty).Value ?? string.Empty,
        };
        if (output.Values.Any(string.IsNullOrEmpty)) return false;
        values = output;
        return true;
    }

    private static bool HasOnlyAttributes(OpenXmlElement source, params string[] names)
    {
        var attributes = source.GetAttributes();
        if (attributes.Count != names.Length) return false;
        var expected = names.ToHashSet(StringComparer.Ordinal);
        return attributes.All(attribute => attribute.NamespaceUri.Length == 0 && expected.Remove(attribute.LocalName)) && expected.Count == 0;
    }

    private static bool HasTransitionAttributes(P.Transition source)
    {
        var attributes = source.GetAttributes();
        if (attributes.Count is < 2 or > 4) return false;
        var allowed = new HashSet<string>(["spd", "advClick", "advTm"], StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var sawDuration = false;
        foreach (var attribute in attributes)
        {
            if (attribute.NamespaceUri == PowerPoint2010Namespace && attribute.LocalName == "dur")
            {
                if (sawDuration) return false;
                sawDuration = true;
                continue;
            }
            if (attribute.NamespaceUri.Length != 0 || !allowed.Contains(attribute.LocalName) || !seen.Add(attribute.LocalName)) return false;
        }
        return seen.Contains("spd") && seen.Contains("advClick");
    }

    private static bool TrySpeed(P.TransitionSpeedValues value, out string name)
    {
        name = value.Equals(P.TransitionSpeedValues.Slow)
            ? "slow"
            : value.Equals(P.TransitionSpeedValues.Medium)
                ? "medium"
                : value.Equals(P.TransitionSpeedValues.Fast)
                    ? "fast"
                    : string.Empty;
        return name.Length > 0;
    }

    private static P.TransitionSpeedValues Speed(string value) => value switch
    {
        "slow" => P.TransitionSpeedValues.Slow,
        "medium" => P.TransitionSpeedValues.Medium,
        "fast" => P.TransitionSpeedValues.Fast,
        _ => throw Invalid("Presentation transition speed is invalid."),
    };

    private static bool TryBooleanToken(string value, out bool result)
    {
        if (value is "1" or "true")
        {
            result = true;
            return true;
        }
        result = false;
        return value is "0" or "false";
    }

    private static bool TryOrientationToken(string value, out string name)
    {
        name = value == "horz" ? "horizontal" : value == "vert" ? "vertical" : string.Empty;
        return name.Length > 0;
    }

    private static string OrientationToken(string value) => value switch
    {
        "horizontal" => "horz",
        "vertical" => "vert",
        _ => throw Invalid("Presentation transition orientation is invalid."),
    };

    private static bool TryDirectionToken(TransitionEffectShape shape, string value, out string name)
    {
        name = value == "l"
            ? "left"
            : value == "u"
                ? "up"
                : value == "r"
                    ? "right"
                    : value == "d"
                        ? "down"
                        : value == "lu"
                            ? "leftUp"
                            : value == "ru"
                                ? "rightUp"
                                : value == "ld"
                                    ? "leftDown"
                                    : value == "rd"
                                        ? "rightDown"
                                        : value is "in" or "out"
                                            ? value
                                            : string.Empty;
        return IsDirection(shape, name);
    }

    private static string DirectionToken(TransitionEffectShape shape, string value)
    {
        if (!IsDirection(shape, value)) throw Invalid("Presentation transition direction is invalid.");
        return value switch
        {
            "left" => "l",
            "up" => "u",
            "right" => "r",
            "down" => "d",
            "leftUp" => "lu",
            "rightUp" => "ru",
            "leftDown" => "ld",
            "rightDown" => "rd",
            "in" or "out" => value,
            _ => throw Invalid("Presentation transition direction is invalid."),
        };
    }

    private static bool IsDirection(TransitionEffectShape shape, string value) => shape switch
    {
        TransitionEffectShape.SlideDirection => value is "left" or "up" or "right" or "down",
        TransitionEffectShape.EightDirection => value is "left" or "up" or "right" or "down" or "leftUp" or "rightUp" or "leftDown" or "rightDown",
        TransitionEffectShape.CornerDirection => value is "leftUp" or "rightUp" or "leftDown" or "rightDown",
        TransitionEffectShape.Split or TransitionEffectShape.InOut => value is "in" or "out",
        _ => false,
    };

    private static string DirectionLabel(TransitionEffectShape shape) => shape switch
    {
        TransitionEffectShape.SlideDirection => "left, up, right, or down",
        TransitionEffectShape.EightDirection => "a cardinal or corner direction",
        TransitionEffectShape.CornerDirection => "leftUp, rightUp, leftDown, or rightDown",
        TransitionEffectShape.Split or TransitionEffectShape.InOut => "in or out",
        _ => "absent",
    };

    private static bool IsOrientation(string value) => value is "horizontal" or "vertical";

    private static void SetAttribute(OpenXmlElement target, string name, string value) =>
        target.SetAttribute(new OpenXmlAttribute(string.Empty, name, string.Empty, value));

    private static bool IsSpeed(string value) => value is "slow" or "medium" or "fast";
    private static bool TryMilliseconds(string value, uint maximum, out uint milliseconds) =>
        uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out milliseconds) && milliseconds <= maximum;
    private static CodecException Invalid(string message) => new("invalid_presentation_transition", message);
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
