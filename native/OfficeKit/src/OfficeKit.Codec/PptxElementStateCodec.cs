using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxElementState(
    bool Hidden,
    bool? Locked,
    bool VisibilityEditable,
    bool LockingEditable);

// PPJ owns one Agent-facing visibility bit and one canonical edit-lock bit.
// DrawingML has a different lock element for every object family, so this
// codec is the single normalization boundary. Partial native lock profiles are
// readable only as source-owned XML and never become an editable boolean.
internal static class PptxElementStateCodec
{
    internal static PptxElementState Read(OpenXmlElement source, PresentationElement element)
    {
        var nonVisual = NonVisual(source);
        var locked = ClassifyLocks(source);
        return new PptxElementState(
            nonVisual?.Hidden?.Value == true,
            locked,
            nonVisual is not null,
            locked is not null);
    }

    internal static void Populate(OpenXmlElement source, PresentationElement element)
    {
        var state = Read(source, element);
        if (state.VisibilityEditable) element.Hidden = state.Hidden;
        if (state.Locked is { } locked) element.Locked = locked;
        if (element.Source is not null)
        {
            element.Source.VisibilityEditable = state.VisibilityEditable;
            element.Source.LockingEditable = state.LockingEditable;
        }
    }

    internal static void ApplyAuthored(OpenXmlElement target, PresentationElement element)
    {
        if (element.HasHidden)
        {
            var nonVisual = NonVisual(target) ??
                throw new CodecException("unsupported_presentation_element_state", $"Presentation element {element.Id} has no non-visual visibility owner.");
            nonVisual.Hidden = element.Hidden ? true : null;
        }
        if (element.HasLocked) ApplyLocks(target, element, element.Locked);
    }

    internal static bool ApplyBound(OpenXmlElement target, PresentationElement original, PresentationElement requested)
    {
        var hiddenChanged = NormalizedHidden(original) != NormalizedHidden(requested);
        var lockedChanged = NormalizedLocked(original) != NormalizedLocked(requested);
        if (!hiddenChanged && !lockedChanged) return false;

        var actual = Read(target, original);
        if (hiddenChanged)
        {
            if (original.Source?.VisibilityEditable != true || !actual.VisibilityEditable ||
                actual.Hidden != NormalizedHidden(original) || !requested.HasHidden)
                throw Unsupported(requested.Id, "hidden state is stale or was not issued as editable");
            NonVisual(target)!.Hidden = requested.Hidden ? true : null;
        }
        if (lockedChanged)
        {
            if (original.Source?.LockingEditable != true || !actual.LockingEditable ||
                actual.Locked != NormalizedLocked(original) || !requested.HasLocked)
                throw Unsupported(requested.Id, "lock state is partial, stale, or was not issued as editable");
            ApplyLocks(target, requested, requested.Locked);
        }
        return true;
    }

    internal static bool EqualExceptState(PresentationElement original, PresentationElement requested)
    {
        var oldSemantic = original.Clone();
        var newSemantic = requested.Clone();
        oldSemantic.ClearHidden();
        oldSemantic.ClearLocked();
        newSemantic.ClearHidden();
        newSemantic.ClearLocked();
        // Source proof and external program identity are validated separately.
        // This comparison decides only whether the native content writer must
        // run after a bounded state mutation.
        oldSemantic.Id = newSemantic.Id = string.Empty;
        oldSemantic.Source = null;
        newSemantic.Source = null;
        var equal = newSemantic.Equals(oldSemantic);
        return equal;
    }

    internal static bool StateChanged(PresentationElement original, PresentationElement requested) =>
        NormalizedHidden(original) != NormalizedHidden(requested) ||
        NormalizedLocked(original) != NormalizedLocked(requested);

    // Residual hashes prove that fields outside the semantic projection did not
    // drift. Remove only state that this codec can fully classify; partial or
    // vendor-specific lock profiles remain in the residual and stay protected.
    internal static void ScrubModeledContent(OpenXmlElement source)
    {
        if (NonVisual(source) is { } nonVisual) nonVisual.Hidden = null;
        if (ClassifyLocks(source) is null) return;

        switch (source)
        {
            case P.Shape shape:
                shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties?.RemoveAllChildren<A.ShapeLocks>();
                return;
            case P.Picture picture:
                picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.RemoveAllChildren<A.PictureLocks>();
                return;
            case P.ConnectionShape connector:
                connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties?.RemoveAllChildren<A.ConnectionShapeLocks>();
                return;
            case P.GraphicFrame frame:
                frame.NonVisualGraphicFrameProperties?.NonVisualGraphicFrameDrawingProperties?.RemoveAllChildren<A.GraphicFrameLocks>();
                return;
            case P.GroupShape group:
                group.NonVisualGroupShapeProperties?.NonVisualGroupShapeDrawingProperties?.RemoveAllChildren<A.GroupShapeLocks>();
                return;
        }
    }

    private static bool NormalizedHidden(PresentationElement element) => element.HasHidden && element.Hidden;
    private static bool NormalizedLocked(PresentationElement element) => element.HasLocked && element.Locked;

    private static P.NonVisualDrawingProperties? NonVisual(OpenXmlElement source) => source switch
    {
        P.Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
        P.Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
        P.ConnectionShape connector => connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
        P.GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
        P.GroupShape group => group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
        _ => null,
    };

    private static bool? ClassifyLocks(OpenXmlElement source) => source switch
    {
        P.Shape shape => ClassifyShapeLocks(shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties),
        P.Picture picture => ClassifyPictureLocks(picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties),
        P.ConnectionShape connector => ClassifyConnectorLocks(connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties),
        P.GraphicFrame frame => ClassifyGraphicFrameLocks(frame.NonVisualGraphicFrameProperties?.NonVisualGraphicFrameDrawingProperties),
        P.GroupShape group => ClassifyGroupLocks(group.NonVisualGroupShapeProperties?.NonVisualGroupShapeDrawingProperties),
        _ => null,
    };

    private static bool? ClassifyShapeLocks(P.NonVisualShapeDrawingProperties? owner)
    {
        if (!TryOnlyLock(owner, out A.ShapeLocks? locks)) return null;
        if (locks is null) return false;
        if (!KnownAttributes(locks, "noGrp", "noSelect", "noRot", "noChangeAspect", "noMove", "noResize", "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType", "noTextEdit")) return null;
        if (AllTrue(locks.NoGrouping, locks.NoSelection, locks.NoRotation, locks.NoChangeAspect, locks.NoMove, locks.NoResize,
                locks.NoEditPoints, locks.NoAdjustHandles, locks.NoChangeArrowheads, locks.NoChangeShapeType, locks.NoTextEdit)) return true;
        return AnyTrue(locks.NoSelection, locks.NoRotation, locks.NoChangeAspect, locks.NoMove, locks.NoResize,
            locks.NoEditPoints, locks.NoAdjustHandles, locks.NoChangeArrowheads, locks.NoChangeShapeType, locks.NoTextEdit) ? null : false;
    }

    private static bool? ClassifyPictureLocks(P.NonVisualPictureDrawingProperties? owner)
    {
        if (!TryOnlyLock(owner, out A.PictureLocks? locks)) return null;
        if (locks is null) return false;
        if (!KnownAttributes(locks, "noGrp", "noSelect", "noRot", "noChangeAspect", "noMove", "noResize", "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType", "noCrop")) return null;
        if (AllTrue(locks.NoGrouping, locks.NoSelection, locks.NoRotation, locks.NoChangeAspect, locks.NoMove, locks.NoResize,
                locks.NoEditPoints, locks.NoAdjustHandles, locks.NoChangeArrowheads, locks.NoChangeShapeType, locks.NoCrop)) return true;
        return AnyTrue(locks.NoSelection, locks.NoRotation, locks.NoMove, locks.NoResize, locks.NoEditPoints,
            locks.NoAdjustHandles, locks.NoChangeArrowheads, locks.NoChangeShapeType, locks.NoCrop) ? null : false;
    }

    private static bool? ClassifyConnectorLocks(P.NonVisualConnectorShapeDrawingProperties? owner)
    {
        if (owner is null) return null;
        A.ConnectionShapeLocks? locks = null;
        foreach (var child in owner.ChildElements)
        {
            if (child is A.StartConnection or A.EndConnection) continue;
            if (child is not A.ConnectionShapeLocks typed || locks is not null) return null;
            locks = typed;
        }
        if (locks is null) return false;
        if (locks.ChildElements.Count != 0) return null;
        if (!KnownAttributes(locks, "noGrp", "noSelect", "noRot", "noChangeAspect", "noMove", "noResize", "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType")) return null;
        if (AllTrue(locks.NoGrouping, locks.NoSelection, locks.NoRotation, locks.NoChangeAspect, locks.NoMove, locks.NoResize,
                locks.NoEditPoints, locks.NoAdjustHandles, locks.NoChangeArrowheads, locks.NoChangeShapeType)) return true;
        return AnyTrue(locks.NoSelection, locks.NoRotation, locks.NoChangeAspect, locks.NoMove, locks.NoResize,
            locks.NoEditPoints, locks.NoAdjustHandles, locks.NoChangeArrowheads, locks.NoChangeShapeType) ? null : false;
    }

    private static bool? ClassifyGraphicFrameLocks(P.NonVisualGraphicFrameDrawingProperties? owner)
    {
        if (!TryOnlyLock(owner, out A.GraphicFrameLocks? locks)) return null;
        if (locks is null) return false;
        if (!KnownAttributes(locks, "noGrp", "noDrilldown", "noSelect", "noChangeAspect", "noMove", "noResize")) return null;
        if (AllTrue(locks.NoGrouping, locks.NoDrilldown, locks.NoSelection, locks.NoChangeAspect, locks.NoMove, locks.NoResize)) return true;
        return AnyTrue(locks.NoDrilldown, locks.NoSelection, locks.NoChangeAspect, locks.NoMove, locks.NoResize) ? null : false;
    }

    private static bool? ClassifyGroupLocks(P.NonVisualGroupShapeDrawingProperties? owner)
    {
        if (!TryOnlyLock(owner, out A.GroupShapeLocks? locks)) return null;
        if (locks is null) return false;
        if (!KnownAttributes(locks, "noGrp", "noUngrp", "noSelect", "noRot", "noChangeAspect", "noMove", "noResize")) return null;
        if (AllTrue(locks.NoGrouping, locks.NoUngrouping, locks.NoSelection, locks.NoRotation, locks.NoChangeAspect, locks.NoMove, locks.NoResize)) return true;
        return AnyTrue(locks.NoUngrouping, locks.NoSelection, locks.NoRotation, locks.NoChangeAspect, locks.NoMove, locks.NoResize) ? null : false;
    }

    internal static bool RecognizesGroupLockProfile(P.NonVisualGroupShapeDrawingProperties owner) =>
        ClassifyGroupLocks(owner) is not null;

    private static bool TryOnlyLock<T>(OpenXmlCompositeElement? owner, out T? locks) where T : OpenXmlElement
    {
        locks = null;
        if (owner is null) return false;
        foreach (var child in owner.ChildElements)
        {
            if (child is not T typed || locks is not null) return false;
            locks = typed;
        }
        return locks?.ChildElements.Count is null or 0;
    }

    private static bool KnownAttributes(OpenXmlElement source, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        foreach (var attribute in source.GetAttributes())
            if (attribute.NamespaceUri.Length > 0 || !allowed.Contains(attribute.LocalName) || !Boolean(attribute.Value ?? string.Empty, out _))
                return false;
        return true;
    }

    private static bool Boolean(string value, out bool result)
    {
        if (value is "1" or "true") { result = true; return true; }
        if (value is "0" or "false") { result = false; return true; }
        result = false;
        return false;
    }

    private static bool AnyTrue(params BooleanValue?[] values) => values.Any(value => value?.Value == true);
    private static bool AllTrue(params BooleanValue?[] values) => values.All(value => value?.Value == true);

    private static void ApplyLocks(OpenXmlElement target, PresentationElement element, bool locked)
    {
        switch (target)
        {
            case P.Shape shape:
                var shapeOwner = shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties ??
                    throw Missing(element.Id);
                shapeOwner.RemoveAllChildren<A.ShapeLocks>();
                if (locked)
                    shapeOwner.Append(new A.ShapeLocks
                    {
                        NoGrouping = true, NoSelection = true, NoRotation = true, NoChangeAspect = true,
                        NoMove = true, NoResize = true, NoEditPoints = true, NoAdjustHandles = true,
                        NoChangeArrowheads = true, NoChangeShapeType = true, NoTextEdit = true,
                    });
                else if (element.Shape?.Placeholder is not null)
                    shapeOwner.Append(new A.ShapeLocks { NoGrouping = true });
                return;
            case P.Picture picture:
                var pictureOwner = picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties ??
                    throw Missing(element.Id);
                pictureOwner.RemoveAllChildren<A.PictureLocks>();
                if (locked)
                    pictureOwner.Append(new A.PictureLocks
                    {
                        NoGrouping = true, NoSelection = true, NoRotation = true, NoChangeAspect = true,
                        NoMove = true, NoResize = true, NoEditPoints = true, NoAdjustHandles = true,
                        NoChangeArrowheads = true, NoChangeShapeType = true, NoCrop = true,
                    });
                else if (element.ContentCase == PresentationElement.ContentOneofCase.Media)
                    pictureOwner.Append(new A.PictureLocks { NoChangeAspect = true });
                return;
            case P.ConnectionShape connector:
                var connectorOwner = connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties ??
                    throw Missing(element.Id);
                connectorOwner.RemoveAllChildren<A.ConnectionShapeLocks>();
                if (locked)
                    connectorOwner.Append(new A.ConnectionShapeLocks
                    {
                        NoGrouping = true, NoSelection = true, NoRotation = true, NoChangeAspect = true,
                        NoMove = true, NoResize = true, NoEditPoints = true, NoAdjustHandles = true,
                        NoChangeArrowheads = true, NoChangeShapeType = true,
                    });
                return;
            case P.GraphicFrame frame:
                var frameOwner = frame.NonVisualGraphicFrameProperties?.NonVisualGraphicFrameDrawingProperties ??
                    throw Missing(element.Id);
                frameOwner.RemoveAllChildren<A.GraphicFrameLocks>();
                frameOwner.Append(locked
                    ? new A.GraphicFrameLocks
                    {
                        NoGrouping = true, NoDrilldown = true, NoSelection = true,
                        NoChangeAspect = true, NoMove = true, NoResize = true,
                    }
                    : new A.GraphicFrameLocks { NoGrouping = true });
                return;
            case P.GroupShape group:
                var groupOwner = group.NonVisualGroupShapeProperties?.NonVisualGroupShapeDrawingProperties ??
                    throw Missing(element.Id);
                groupOwner.RemoveAllChildren<A.GroupShapeLocks>();
                if (locked)
                    groupOwner.Append(new A.GroupShapeLocks
                    {
                        NoGrouping = true, NoUngrouping = true, NoSelection = true, NoRotation = true,
                        NoChangeAspect = true, NoMove = true, NoResize = true,
                    });
                return;
            default:
                throw Missing(element.Id);
        }
    }

    private static CodecException Missing(string id) => new(
        "unsupported_presentation_element_state",
        $"Presentation element {id} has no canonical non-visual lock owner.");

    private static CodecException Unsupported(string id, string detail) => new(
        "unsupported_presentation_element_state",
        $"Presentation element {id} {detail}.");
}
