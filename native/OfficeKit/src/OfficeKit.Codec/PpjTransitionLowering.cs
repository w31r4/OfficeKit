using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// PPJ exposes the same bounded semantic transition profile as the Presentation
// wire. Defaults live here so authored and source-bound compilation cannot
// drift while the native codec remains the final validity oracle.
internal static class PpjTransitionLowering
{
    private sealed record Profile(
        string? DefaultDirection = null,
        string? DefaultOrientation = null,
        bool AllowsThroughBlack = false,
        int? DefaultSpokes = null);

    private static readonly IReadOnlyDictionary<string, Profile> Profiles =
        new Dictionary<string, Profile>(StringComparer.Ordinal)
        {
            ["blinds"] = new(DefaultOrientation: "horizontal"),
            ["checker"] = new(DefaultOrientation: "horizontal"),
            ["circle"] = new(),
            ["comb"] = new(DefaultOrientation: "horizontal"),
            ["cover"] = new(DefaultDirection: "left"),
            ["cut"] = new(AllowsThroughBlack: true),
            ["diamond"] = new(),
            ["dissolve"] = new(),
            ["fade"] = new(AllowsThroughBlack: true),
            ["newsflash"] = new(),
            ["plus"] = new(),
            ["pull"] = new(DefaultDirection: "left"),
            ["push"] = new(DefaultDirection: "left"),
            ["random"] = new(),
            ["randomBar"] = new(DefaultOrientation: "horizontal"),
            ["split"] = new(DefaultDirection: "out", DefaultOrientation: "vertical"),
            ["strips"] = new(DefaultDirection: "rightDown"),
            ["wedge"] = new(),
            ["wheel"] = new(DefaultSpokes: 1),
            ["wipe"] = new(DefaultDirection: "left"),
            ["zoom"] = new(DefaultDirection: "in"),
        };

    internal static bool IsBaseEffect(string type) => Profiles.ContainsKey(type);

    internal static bool TryBuildBase(
        PpjTransitionModel source,
        out PresentationTransition transition,
        out string? error)
    {
        transition = new PresentationTransition();
        error = null;
        if (!Profiles.TryGetValue(source.Type, out var profile))
        {
            error = $"PPJ transition type {source.Type} is not a base transition.";
            return false;
        }

        if (source.ThroughBlack is not null && !profile.AllowsThroughBlack)
        {
            error = $"PPJ {source.Type} transition does not accept throughBlack.";
            return false;
        }
        if (source.Spokes is not null && profile.DefaultSpokes is null)
        {
            error = $"PPJ {source.Type} transition does not accept spokes.";
            return false;
        }
        if (source.Direction is not null && profile.DefaultDirection is null)
        {
            error = $"PPJ {source.Type} transition does not accept direction.";
            return false;
        }
        if (source.Orientation is not null && profile.DefaultOrientation is null)
        {
            error = $"PPJ {source.Type} transition does not accept orientation.";
            return false;
        }

        transition.Effect = source.Type;
        transition.Speed = source.Speed ?? "medium";
        transition.AdvanceOnClick = source.AdvanceOnClick ?? true;
        if ((source.Direction ?? profile.DefaultDirection) is { } direction)
            transition.Direction = direction;
        if ((source.Orientation ?? profile.DefaultOrientation) is { } orientation)
            transition.Orientation = orientation;
        if (source.ThroughBlack is { } throughBlack)
            transition.ThroughBlack = throughBlack;
        if ((source.Spokes ?? profile.DefaultSpokes) is { } spokes)
            transition.Spokes = checked((uint)spokes);
        if (source.DurationMs is { } durationMs)
            transition.DurationMs = checked((uint)durationMs);
        if (source.AdvanceAfterMs is { } advanceAfterMs)
            transition.AdvanceAfterMs = checked((uint)advanceAfterMs);

        try
        {
            PptxTransitionCodec.Validate(transition);
            return true;
        }
        catch (CodecException exception)
        {
            error = exception.Message;
            transition = new PresentationTransition();
            return false;
        }
    }

    internal static PresentationTransition BuildBase(PpjTransitionModel source)
    {
        if (TryBuildBase(source, out var transition, out var error)) return transition;
        throw new CodecException("invalid_ppj_transition", error ?? "PPJ base transition is invalid.");
    }
}
