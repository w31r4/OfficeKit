using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Tracks relationships rooted at one PresentationML owner. Asset identity
// belongs to the shared catalog; relationship IDs remain local to their actual
// source part (including a nested SmartArt cached drawing).
internal sealed class PptxPartContext
{
    private const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    private readonly HashSet<string> _addedRelationshipIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _addedRelationshipKeys = new(StringComparer.Ordinal);
    private readonly HashSet<string> _addedPartPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _removedRelationshipIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _removedRelationshipKeys = new(StringComparer.Ordinal);
    private readonly HashSet<string> _removedPartPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly Func<PartTypeInfo, string, ImagePart> _addImagePart;

    internal PptxPartContext(
        OpenXmlPart owner,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        IReadOnlyDictionary<string, SlidePart>? slidePartById = null,
        PptxAssetCatalog? assets = null,
        PptxCustomShowCatalog? customShows = null,
        int? slideNumber = null,
        bool deriveAutomaticFields = false) : this(
            owner,
            owner switch
            {
                SlidePart slide => (type, relationshipId) => slide.AddImagePart(type, relationshipId),
                SlideMasterPart master => (type, relationshipId) => master.AddImagePart(type, relationshipId),
                SlideLayoutPart layout => (type, relationshipId) => layout.AddImagePart(type, relationshipId),
                ChartPart chart => (type, relationshipId) => chart.AddImagePart(type, relationshipId),
                DiagramPersistLayoutPart drawing => (type, relationshipId) => drawing.AddImagePart(type, relationshipId),
                // NotesSlide rich text deliberately has no relationship-writing
                // surface. The context exists so the shared text codec can
                // preserve its fixed paragraph/run topology; a picture bullet
                // still fails closed before it can add an ImagePart.
                NotesSlidePart => (_, _) => throw new CodecException(
                    "unsupported_presentation_notes",
                    "Speaker notes cannot add picture-bullet relationships."),
                _ => throw new ArgumentException($"Unsupported PresentationML relationship owner {owner.GetType().Name}.", nameof(owner)),
            },
            slideIdByPartPath,
            slidePartById,
            assets,
            customShows,
            slideNumber,
            deriveAutomaticFields)
    {
    }

    private PptxPartContext(
        OpenXmlPart owner,
        Func<PartTypeInfo, string, ImagePart> addImagePart,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        IReadOnlyDictionary<string, SlidePart>? slidePartById,
        PptxAssetCatalog? assets,
        PptxCustomShowCatalog? customShows,
        int? slideNumber,
        bool deriveAutomaticFields)
    {
        Owner = owner;
        _addImagePart = addImagePart;
        SlideIdByPartPath = slideIdByPartPath;
        SlidePartById = slidePartById ?? new Dictionary<string, SlidePart>(StringComparer.Ordinal);
        Assets = assets;
        CustomShows = customShows ?? PptxCustomShowCatalog.Empty;
        SlideNumber = slideNumber;
        DeriveAutomaticFields = deriveAutomaticFields;
    }

    internal OpenXmlPart Owner { get; }
    internal IReadOnlyDictionary<string, string> SlideIdByPartPath { get; }
    internal IReadOnlyDictionary<string, SlidePart> SlidePartById { get; }
    internal PptxAssetCatalog? Assets { get; }
    internal PptxCustomShowCatalog CustomShows { get; }
    internal int? SlideNumber { get; }
    internal bool DeriveAutomaticFields { get; set; }
    internal bool RelationshipsChanged => _addedRelationshipKeys.Count > 0 || _removedRelationshipKeys.Count > 0;
    internal IReadOnlyCollection<string> AddedRelationshipIds => _addedRelationshipIds;
    internal IReadOnlyCollection<string> AddedRelationshipKeys => _addedRelationshipKeys;
    internal IReadOnlyCollection<string> AddedPartPaths => _addedPartPaths;
    internal IReadOnlyCollection<string> RemovedRelationshipIds => _removedRelationshipIds;
    internal IReadOnlyCollection<string> RemovedRelationshipKeys => _removedRelationshipKeys;
    internal IReadOnlyCollection<string> RemovedPartPaths => _removedPartPaths;

    internal void TrackAddedPart(OpenXmlPart part)
    {
        TrackAddedPart(Owner, part);
    }

    // A SmartArt cached drawing owns its image relationships. Keep the
    // relationship source path with the slide context so source-preserving
    // export can authorize the nested .rels part without flattening it onto
    // the slide relationship set.
    internal void TrackAddedPart(OpenXmlPart relationshipOwner, OpenXmlPart part)
    {
        var relationshipId = relationshipOwner.GetIdOfPart(part);
        TrackAddedRelationship(relationshipOwner, relationshipId);
        _addedPartPaths.Add(part.Uri.OriginalString.TrimStart('/'));
    }

    internal string NextRelationshipId(string stem)
    {
        var used = Owner.Parts.Select(pair => pair.RelationshipId)
            .Concat(Owner.ExternalRelationships.Select(relationship => relationship.Id))
            .Concat(Owner.HyperlinkRelationships.Select(relationship => relationship.Id))
            .Concat(Owner.DataPartReferenceRelationships.Select(relationship => relationship.Id))
            .Concat(_addedRelationshipIds)
            .ToHashSet(StringComparer.Ordinal);
        for (var index = 1; index <= 1_000_000; index++)
        {
            var candidate = index == 1 ? stem : $"{stem}_{index}";
            if (!used.Contains(candidate)) return candidate;
        }
        throw new CodecException(
            "presentation_relationship_budget_exceeded",
            "PPTX relationship ID allocation exceeded its bounded search.");
    }

    internal string AddExternalHyperlink(string uri)
    {
        var existing = Owner.HyperlinkRelationships.FirstOrDefault(relationship =>
            relationship.IsExternal && relationship.Uri.OriginalString.Equals(uri, StringComparison.Ordinal));
        if (existing is not null) return existing.Id;
        return Track(Owner.AddHyperlinkRelationship(new Uri(uri, UriKind.Absolute), true).Id);
    }

    internal string AddSlide(string slideId)
    {
        if (!SlidePartById.TryGetValue(slideId, out var target))
            throw new CodecException("invalid_presentation_hyperlink", $"Presentation run hyperlink references missing slide {slideId}.");
        var existing = Owner.Parts.FirstOrDefault(pair => ReferenceEquals(pair.OpenXmlPart, target));
        if (existing.OpenXmlPart is not null) return existing.RelationshipId;
        Owner.AddPart(target);
        return Track(Owner.GetIdOfPart(target));
    }

    internal bool TryReadPicture(A.PictureBullet source, out PresentationPictureBullet picture)
    {
        picture = new PresentationPictureBullet();
        if (Assets is null || source.ChildElements.Count != 1 || source.GetFirstChild<A.Blip>() is not { } blip ||
            blip.ChildElements.Count > 0 || blip.CompressionState is not null) return false;
        var embed = blip.Embed?.Value ?? string.Empty;
        var link = blip.Link?.Value ?? string.Empty;
        if ((embed.Length == 0) == (link.Length == 0)) return false;
        if (embed.Length > 0)
        {
            try
            {
                if (Owner.GetPartById(embed) is not ImagePart imagePart) return false;
                picture.AssetId = Assets.Import(imagePart).Id;
                return true;
            }
            catch (Exception error) when (error is ArgumentOutOfRangeException or CodecException)
            {
                return false;
            }
        }
        var relationship = Owner.ExternalRelationships.FirstOrDefault(item => item.Id == link && item.RelationshipType.EndsWith("/image", StringComparison.Ordinal));
        if (relationship is null) return false;
        try
        {
            picture.Uri = ValidatePictureUri(relationship.Uri.OriginalString);
            return true;
        }
        catch (CodecException)
        {
            picture = new PresentationPictureBullet();
            return false;
        }
    }

    internal bool IsPictureRelationship(A.PictureBullet source)
    {
        if (source.ChildElements.Count != 1 || source.GetFirstChild<A.Blip>() is not { } blip ||
            blip.ChildElements.Count > 0 || blip.CompressionState is not null)
            return false;
        var embed = blip.Embed?.Value ?? string.Empty;
        var link = blip.Link?.Value ?? string.Empty;
        if ((embed.Length == 0) == (link.Length == 0)) return false;
        return embed.Length > 0
            ? Owner.GetPartById(embed) is ImagePart
            : Owner.ExternalRelationships.Any(relationship =>
                relationship.Id == link && relationship.RelationshipType.EndsWith("/image", StringComparison.Ordinal));
    }

    internal A.PictureBullet BuildPicture(PresentationPictureBullet picture)
    {
        var blip = new A.Blip();
        switch (picture.SourceCase)
        {
            case PresentationPictureBullet.SourceOneofCase.AssetId:
                blip.Embed = AddEmbeddedPicture(picture.AssetId);
                break;
            case PresentationPictureBullet.SourceOneofCase.Uri:
                blip.Link = AddExternalPicture(picture.Uri);
                break;
            default:
                throw InvalidPicture("Presentation picture bullet requires exactly one source.");
        }
        return new A.PictureBullet(blip);
    }

    internal static void ValidatePicture(PresentationPictureBullet? picture)
    {
        if (picture is null) throw InvalidPicture("Presentation picture bullet payload is missing.");
        switch (picture.SourceCase)
        {
            case PresentationPictureBullet.SourceOneofCase.AssetId:
                if (string.IsNullOrWhiteSpace(picture.AssetId) || picture.AssetId.Length > 512)
                    throw InvalidPicture("Presentation picture bullet asset ID must contain 1 through 512 characters.");
                break;
            case PresentationPictureBullet.SourceOneofCase.Uri:
                ValidatePictureUri(picture.Uri);
                break;
            default:
                throw InvalidPicture("Presentation picture bullet requires exactly one source.");
        }
    }

    internal Asset ReadEmbeddedPicture(string relationshipId)
    {
        if (Assets is null) throw InvalidPicture("Presentation image import requires an asset catalog.");
        try
        {
            if (Owner.GetPartById(relationshipId) is not ImagePart imagePart)
                throw InvalidPicture($"Presentation image relationship {relationshipId} does not resolve to an image part.");
            return Assets.Import(imagePart);
        }
        catch (ArgumentOutOfRangeException)
        {
            throw InvalidPicture($"Presentation image relationship {relationshipId} is missing.");
        }
    }

    internal string AddEmbeddedPicture(string assetId)
    {
        if (Assets is null) throw InvalidPicture("Presentation picture authoring requires an asset catalog.");
        var asset = Assets.Get(assetId);
        var existingOwnerPart = Owner.Parts.Select(pair => pair.OpenXmlPart).OfType<ImagePart>().FirstOrDefault(part => PartMatches(part, asset));
        if (existingOwnerPart is not null)
        {
            Assets.RegisterPart(assetId, existingOwnerPart);
            return Owner.GetIdOfPart(existingOwnerPart);
        }
        if (Assets.ExistingPart(assetId) is { } shared)
        {
            var sharedRelationshipId = NextRelationshipId(asset);
            Owner.AddPart(shared, sharedRelationshipId);
            return Track(sharedRelationshipId);
        }
        if (asset.Data.IsEmpty)
            throw InvalidPicture($"Presentation source package does not contain image asset {assetId}.");
        var relationshipId = NextRelationshipId(asset);
        var part = _addImagePart(PptxAssetCatalog.ImagePartTypeFor(asset.ContentType), relationshipId);
        using (var source = new MemoryStream(asset.Data.ToByteArray(), writable: false)) part.FeedData(source);
        Assets.RegisterPart(assetId, part);
        _addedPartPaths.Add(part.Uri.OriginalString.TrimStart('/'));
        return Track(relationshipId);
    }

    // Removing a modeled image reference must not leave an orphaned source
    // relationship or media part in the owner part. Keep the relationship when
    // another blip in this owner still uses it; a shared image used by another
    // owner remains managed by the package graph.
    internal void RemoveIfUnreferenced(string relationshipId)
    {
        if (string.IsNullOrWhiteSpace(relationshipId) ||
            Owner.RootElement is { } root &&
            (root.GetAttributes().Any(attribute => attribute.Value == relationshipId) ||
             root.Descendants().Any(element => element.GetAttributes().Any(attribute => attribute.Value == relationshipId))))
            return;
        var part = Owner.GetPartById(relationshipId);
        if (part is null) return;
        var partPath = part.Uri.OriginalString.TrimStart('/');
        if (Owner.DeletePart(part))
        {
            _removedRelationshipIds.Add(relationshipId);
            _removedRelationshipKeys.Add(RelationshipKey(Owner, relationshipId));
            _removedPartPaths.Add(partPath);
        }
    }

    // Action-setting replacement owns either a hyperlink relationship or a
    // relationship to another SlidePart.  Remove only that relationship when
    // the old click target is no longer referenced; never delete the target
    // SlidePart itself because it is still owned by ppt/presentation.xml.
    internal void RemoveElementActionRelationshipIfUnreferenced(string relationshipId)
    {
        if (string.IsNullOrWhiteSpace(relationshipId) ||
            Owner.RootElement is { } root &&
            (root.GetAttributes().Any(attribute => attribute.Value == relationshipId) ||
             root.Descendants().Any(element => element.GetAttributes().Any(attribute => attribute.Value == relationshipId))))
            return;

        var hyperlink = Owner.HyperlinkRelationships.FirstOrDefault(item => item.Id == relationshipId);
        if (hyperlink is not null)
        {
            Owner.DeleteReferenceRelationship(relationshipId);
            _removedRelationshipIds.Add(relationshipId);
            _removedRelationshipKeys.Add(RelationshipKey(Owner, relationshipId));
            return;
        }

        var partRelationship = Owner.Parts.FirstOrDefault(item => item.RelationshipId == relationshipId);
        if (partRelationship.OpenXmlPart is SlidePart)
        {
            Owner.DeleteReferenceRelationship(relationshipId);
            _removedRelationshipIds.Add(relationshipId);
            _removedRelationshipKeys.Add(RelationshipKey(Owner, relationshipId));
        }
    }

    private string AddExternalPicture(string value)
    {
        var uri = ValidatePictureUri(value);
        var existing = Owner.ExternalRelationships.FirstOrDefault(relationship =>
            relationship.RelationshipType.EndsWith("/image", StringComparison.Ordinal) &&
            relationship.Uri.OriginalString.Equals(uri, StringComparison.Ordinal));
        if (existing is not null) return existing.Id;
        return Track(Owner.AddExternalRelationship(ImageRelationshipType, new Uri(uri, UriKind.Absolute)).Id);
    }

    private static bool PartMatches(ImagePart part, Asset asset)
    {
        if (!part.ContentType.Equals(asset.ContentType, StringComparison.OrdinalIgnoreCase)) return false;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var sha = System.Security.Cryptography.SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream)).Equals(asset.Sha256, StringComparison.OrdinalIgnoreCase);
    }

    private string Track(string relationshipId)
    {
        TrackAddedRelationship(Owner, relationshipId);
        return relationshipId;
    }

    private void TrackAddedRelationship(OpenXmlPart relationshipOwner, string relationshipId)
    {
        if (ReferenceEquals(relationshipOwner, Owner))
            _addedRelationshipIds.Add(relationshipId);
        _addedRelationshipKeys.Add(RelationshipKey(relationshipOwner, relationshipId));
    }

    private static string RelationshipKey(OpenXmlPart relationshipOwner, string relationshipId) =>
        $"{relationshipOwner.Uri.OriginalString.TrimStart('/')}\0{relationshipId}";

    private string NextRelationshipId(Asset asset)
    {
        var used = Owner.Parts.Select(pair => pair.RelationshipId)
            .Concat(Owner.ExternalRelationships.Select(relationship => relationship.Id))
            .Concat(Owner.HyperlinkRelationships.Select(relationship => relationship.Id))
            .Concat(Owner.DataPartReferenceRelationships.Select(relationship => relationship.Id))
            .Concat(_addedRelationshipIds)
            .ToHashSet(StringComparer.Ordinal);
        var digest = asset.Sha256.Length >= 16 ? asset.Sha256[..16].ToLowerInvariant() : asset.Sha256.ToLowerInvariant();
        var stem = $"rIdOfficeKitImage{digest}_";
        for (var index = 1; index <= 1_000_000; index++)
        {
            var candidate = stem + index;
            if (!used.Contains(candidate)) return candidate;
        }
        throw new CodecException(
            "presentation_relationship_budget_exceeded",
            "PPTX image relationship ID allocation exceeded its bounded search.");
    }

    private static string ValidatePictureUri(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 4_096 || value.Any(char.IsControl) ||
            !Uri.TryCreate(value, UriKind.Absolute, out var uri) ||
            uri.Scheme is not ("http" or "https"))
            throw InvalidPicture("Presentation picture bullet URI must be an absolute http(s) URI of at most 4096 characters without controls.");
        return value;
    }

    private static CodecException InvalidPicture(string message) => new("invalid_presentation_asset", message);
}
