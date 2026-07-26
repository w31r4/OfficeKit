using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxViewPropertiesChange(string PartPath, string Sha256);

// Owns one deliberately narrow imported view-properties edit profile. The
// file-level grid spacing, snap flags, and existing guide positions are useful
// authoring aids, but they must not become permission to rebuild a PowerPoint
// view graph or to persist the JavaScript-only editor visibility toggles.
internal static class PptxViewPropertiesCodec
{
    private const int MaxGuides = 1_024;

    internal static PresentationViewProperties? Read(PresentationPart owner)
    {
        var part = owner.ViewPropertiesPart;
        if (part is null) return null;
        var root = part.ViewProperties ??
            throw new CodecException("missing_presentation_view_root", "PPTX view-properties part has no p:viewPr root.", PartPath(part));
        var result = ReadSemantic(root);
        result.Source = new PresentationViewPropertiesSourceBinding
        {
            PartPath = PartPath(part),
            RelationshipId = owner.GetIdOfPart(part),
            ViewXmlSha256 = HashElement(root),
            SemanticSha256 = SemanticHash(result),
            ResidualSha256 = ResidualHash(root),
            Editable = Supports(part, root),
        };
        return result;
    }

    // Re-proves the original package before applying a fixed-topology semantic
    // delta. The returned hash is the exact rewritten part payload that the
    // generic opaque-OPC guard must permit; relationships and every other
    // opaque part remain fixed.
    internal static PptxViewPropertiesChange? ApplySourceBound(PresentationPart owner, PresentationViewProperties? requested)
    {
        var actual = Read(owner);
        if (actual is null && requested is null) return null;
        if (actual is null || requested?.Source is null)
            throw new CodecException(
                "presentation_view_topology_changed",
                "Source-preserving PPTX export requires the original presentation view-properties topology.",
                "ppt/presentation.xml");

        var part = owner.ViewPropertiesPart ??
            throw new CodecException("missing_presentation_view_part", "PPTX presentation view-properties part disappeared before export.", "ppt/presentation.xml");
        var root = part.ViewProperties ??
            throw new CodecException("missing_presentation_view_root", "PPTX view-properties part has no p:viewPr root.", PartPath(part));
        var binding = requested.Source;
        var source = actual.Source ??
            throw new CodecException("presentation_view_source_binding_mismatch", "Presentation view properties are missing their source binding.", PartPath(part));
        if (!BindingEquals(binding, source))
            throw new CodecException(
                "presentation_view_source_binding_mismatch",
                "Presentation view properties no longer match their hash-bound source part.",
                source.PartPath);

        if (SemanticHash(requested).Equals(SemanticHash(actual), StringComparison.OrdinalIgnoreCase)) return null;
        if (!binding.Editable || !Supports(part, root))
            throw new CodecException(
                "unsupported_presentation_view_edit",
                "Imported presentation grid spacing, snap settings, and guides are preserved but not safely editable by this fixed-topology codec profile.",
                source.PartPath);
        if (!TopologyEquals(actual, requested))
            throw new CodecException(
                "presentation_view_topology_changed",
                "Imported presentation view properties keep grid/snap attribute presence plus guide count, order, and orientation fixed.",
                source.PartPath);
        if (!ResidualHash(root).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_view_source_binding_mismatch",
                "Presentation view properties no longer preserve the bound non-editable XML residual.",
                source.PartPath);

        var residualBefore = ResidualHash(root);
        Apply(root, requested);
        if (!ResidualHash(root).Equals(residualBefore, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_view_residual_not_preserved",
                "Presentation view edit changed non-editable source XML.",
                source.PartPath);
        root.Save();

        var output = Read(owner) ??
            throw new CodecException("presentation_view_semantics_not_applied", "Presentation view properties disappeared after the requested edit.", source.PartPath);
        if (!SemanticHash(output).Equals(SemanticHash(requested), StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_view_semantics_not_applied",
                "Presentation view edit did not produce the requested grid, snap, and guide semantics.",
                source.PartPath);
        return new PptxViewPropertiesChange(PartPath(part), HashPart(part));
    }

    internal static void Validate(PresentationViewProperties? properties, bool hasSourcePackage)
    {
        if (properties is null) return;
        if (!hasSourcePackage || properties.Source is null)
            throw new CodecException(
                "unsupported_presentation_features",
                "Source-free authoring of PowerPoint view properties and guides is unsupported; use presentation.view for local editor visibility.");
        if (string.IsNullOrWhiteSpace(properties.Source.PartPath) ||
            string.IsNullOrWhiteSpace(properties.Source.RelationshipId) ||
            string.IsNullOrWhiteSpace(properties.Source.ViewXmlSha256) ||
            string.IsNullOrWhiteSpace(properties.Source.SemanticSha256) ||
            string.IsNullOrWhiteSpace(properties.Source.ResidualSha256))
            throw new CodecException("invalid_presentation_view", "Presentation view source binding is incomplete.");
        if ((properties.HasGridSpacingCxEmu && properties.GridSpacingCxEmu is <= 0 or > int.MaxValue) ||
            (properties.HasGridSpacingCyEmu && properties.GridSpacingCyEmu is <= 0 or > int.MaxValue))
            throw new CodecException("invalid_presentation_view", "Presentation grid spacing must fit the positive signed 32-bit EMU range when present.");
        foreach (var guide in properties.SlideGuides)
            if (guide.Orientation is not PresentationSlideGuide.Types.Orientation.Horizontal and
                not PresentationSlideGuide.Types.Orientation.Vertical)
                throw new CodecException("invalid_presentation_view", "Presentation guides require horizontal or vertical orientation.");
        if (properties.SlideGuides.Count > MaxGuides)
            throw new CodecException("presentation_guide_budget_exceeded", $"Presentation cannot contain more than {MaxGuides} guides.");
    }

    private static PresentationViewProperties ReadSemantic(P.ViewProperties root)
    {
        var result = new PresentationViewProperties();
        if (root.GridSpacing?.Cx?.Value is { } cx) result.GridSpacingCxEmu = cx;
        if (root.GridSpacing?.Cy?.Value is { } cy) result.GridSpacingCyEmu = cy;
        var common = root.SlideViewProperties?.CommonSlideViewProperties;
        if (common?.SnapToGrid?.Value is { } snapToGrid) result.SlideViewSnapToGrid = snapToGrid;
        if (common?.SnapToObjects?.Value is { } snapToObjects) result.SlideViewSnapToObjects = snapToObjects;
        foreach (var guide in common?.GuideList?.Elements<P.Guide>() ?? [])
        {
            if (result.SlideGuides.Count >= MaxGuides)
                throw new CodecException(
                    "presentation_guide_budget_exceeded",
                    $"PPTX presentation view exceeds the {MaxGuides}-guide budget.");
            if (guide.Position?.Value is not { } position) continue;
            result.SlideGuides.Add(new PresentationSlideGuide
            {
                Orientation = guide.Orientation?.Value == P.DirectionValues.Vertical
                    ? PresentationSlideGuide.Types.Orientation.Vertical
                    : PresentationSlideGuide.Types.Orientation.Horizontal,
                Position = position,
            });
        }
        return result;
    }

    private static bool Supports(ViewPropertiesPart part, P.ViewProperties root)
    {
        if (part.Parts.Any() || part.ExternalRelationships.Any() || part.HyperlinkRelationships.Any() ||
            part.DataPartReferenceRelationships.Any() || root.ExtendedAttributes.Any() ||
            root.ChildElements.Any(child => child.LocalName == "extLst")) return false;
        var slideViews = root.Elements<P.SlideViewProperties>().ToArray();
        var gridSpacings = root.Elements<P.GridSpacing>().ToArray();
        if (slideViews.Length != 1 || gridSpacings.Length > 1) return false;
        if (gridSpacings.Length == 1 &&
            (gridSpacings[0].ExtendedAttributes.Any() || gridSpacings[0].Cx?.Value is null || gridSpacings[0].Cy?.Value is null)) return false;
        var slideView = slideViews[0];
        var common = slideView.CommonSlideViewProperties;
        if (slideView.ExtendedAttributes.Any() || common is null || common.ExtendedAttributes.Any() ||
            slideView.ChildElements.Count != 1 || slideView.ChildElements[0] != common ||
            common.ChildElements.Any(child => child.LocalName is not "cViewPr" and not "guideLst")) return false;
        if (common.ChildElements.Count(child => child.LocalName == "cViewPr") != 1 ||
            common.ChildElements.Count(child => child.LocalName == "guideLst") > 1) return false;
        var guideLists = common.Elements<P.GuideList>().ToArray();
        if (guideLists.Length > 1) return false;
        var guides = guideLists.SingleOrDefault();
        if (guides is null) return true;
        if (guides.ExtendedAttributes.Any() || guides.ChildElements.Count > MaxGuides ||
            guides.ChildElements.Any(child => child is not P.Guide)) return false;
        return guides.Elements<P.Guide>().All(guide =>
        {
            var orientation = guide.Orientation?.Value;
            return !guide.ExtendedAttributes.Any() && guide.ChildElements.Count == 0 &&
                   guide.Position?.Value is not null &&
                   (orientation == P.DirectionValues.Horizontal || orientation == P.DirectionValues.Vertical);
        });
    }

    private static bool TopologyEquals(PresentationViewProperties actual, PresentationViewProperties requested)
    {
        if (actual.HasGridSpacingCxEmu != requested.HasGridSpacingCxEmu ||
            actual.HasGridSpacingCyEmu != requested.HasGridSpacingCyEmu ||
            actual.HasSlideViewSnapToGrid != requested.HasSlideViewSnapToGrid ||
            actual.HasSlideViewSnapToObjects != requested.HasSlideViewSnapToObjects ||
            actual.SlideGuides.Count != requested.SlideGuides.Count) return false;
        return actual.SlideGuides.Zip(requested.SlideGuides).All(pair => pair.First.Orientation == pair.Second.Orientation);
    }

    private static void Apply(P.ViewProperties root, PresentationViewProperties requested)
    {
        var grid = root.GridSpacing;
        if (requested.HasGridSpacingCxEmu) grid!.Cx = checked((int)requested.GridSpacingCxEmu);
        if (requested.HasGridSpacingCyEmu) grid!.Cy = checked((int)requested.GridSpacingCyEmu);
        var common = root.SlideViewProperties!.CommonSlideViewProperties!;
        if (requested.HasSlideViewSnapToGrid) common.SnapToGrid = requested.SlideViewSnapToGrid;
        if (requested.HasSlideViewSnapToObjects) common.SnapToObjects = requested.SlideViewSnapToObjects;
        var guides = common.GuideList?.Elements<P.Guide>().ToArray() ?? [];
        for (var index = 0; index < guides.Length; index++) guides[index].Position = requested.SlideGuides[index].Position;
    }

    private static string ResidualHash(P.ViewProperties root)
    {
        var clone = (P.ViewProperties)root.CloneNode(true);
        if (clone.GridSpacing is { } grid)
        {
            grid.Cx = null;
            grid.Cy = null;
        }
        var common = clone.SlideViewProperties?.CommonSlideViewProperties;
        if (common is not null)
        {
            common.SnapToGrid = null;
            common.SnapToObjects = null;
            foreach (var guide in common.GuideList?.Elements<P.Guide>() ?? []) guide.Position = null;
        }
        return HashElement(clone);
    }

    private static bool BindingEquals(PresentationViewPropertiesSourceBinding left, PresentationViewPropertiesSourceBinding right) =>
        left.PartPath.Equals(right.PartPath, StringComparison.OrdinalIgnoreCase) &&
        left.RelationshipId.Equals(right.RelationshipId, StringComparison.Ordinal) &&
        left.ViewXmlSha256.Equals(right.ViewXmlSha256, StringComparison.OrdinalIgnoreCase) &&
        left.SemanticSha256.Equals(right.SemanticSha256, StringComparison.OrdinalIgnoreCase) &&
        left.ResidualSha256.Equals(right.ResidualSha256, StringComparison.OrdinalIgnoreCase) &&
        left.Editable == right.Editable;

    private static string SemanticHash(PresentationViewProperties properties)
    {
        var semantic = properties.Clone();
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string HashPart(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var bytes = new MemoryStream();
        stream.CopyTo(bytes);
        return Hash(bytes.GetBuffer().AsSpan(0, checked((int)bytes.Length)));
    }
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
