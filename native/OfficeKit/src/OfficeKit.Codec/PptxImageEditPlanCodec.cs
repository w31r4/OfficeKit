using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxImageEditPlanProof(
    Asset Replacement,
    string SourceRelationshipId,
    string SourceImagePartPath,
    string RelationshipPartPath,
    string RelationshipType,
    string ReplacementRelationshipId,
    string ReplacementPartPath,
    string ReplacementTarget,
    bool AddsPart);

// Keeps an imported picture replacement inside the finite Edit Plan IR. The
// slide token stream is patched in place, the old relationship/media graph is
// retained, and one content-addressed copy-on-write image edge is appended.
internal static partial class PptxEditPlanCodec
{
    private const string OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictOfficeRelationshipsNamespace = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private const string PictureAssetPrefix = "asset/presentation/picture-bullet/";
    private const string SvgBlipNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    private static void ValidateImageReplacement(PresentationEditOperation operation)
    {
        var replacement = operation.ImageReplacement;
        var assetId = LeafKind(operation) == "imageSvgAsset" ? replacement?.SvgAssetId : replacement?.AssetId;
        if (replacement is null || string.IsNullOrWhiteSpace(assetId) || assetId.Length > 512)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX image operation {operation.OperationId} has an invalid replacement asset ID.");
        if (replacement.Crop is { } crop && !ValidImageCrop(crop))
            throw new CodecException("invalid_presentation_edit_operation", $"PPTX image operation {operation.OperationId} has an invalid source rectangle.");
    }

    private static PptxImageEditPlanProof ProveImageReplacement(
        byte[] sourceBytes,
        SlidePart slidePart,
        OpenXmlElement element,
        PresentationElement projectedElement,
        PresentationEditOperation operation,
        PptxAssetCatalog requestedAssets)
    {
        if (element is not P.Picture picture || projectedElement.ContentCase != PresentationElement.ContentOneofCase.Image ||
            projectedElement.Source.Editable != true || projectedElement.Image.AssetId != operation.ExpectedValue)
            throw new CodecException("unsupported_presentation_edit", $"PPTX image operation {operation.OperationId} target is not a bounded editable picture.", operation.SlidePartPath);
        var relationshipId = picture.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value ?? string.Empty;
        if (relationshipId.Length == 0)
            throw new CodecException("presentation_edit_target_missing", $"PPTX image operation {operation.OperationId} source relationship is missing.", operation.SlidePartPath);
        ImagePart sourcePart;
        try
        {
            sourcePart = slidePart.GetPartById(relationshipId!) as ImagePart ??
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} relationship is not an ImagePart.", operation.SlidePartPath);
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw new CodecException("presentation_edit_target_missing", $"PPTX image operation {operation.OperationId} source relationship is missing.", operation.SlidePartPath, exception);
        }
        var sourceData = ReadOpenXmlPart(sourcePart);
        var sourceAssetId = PictureAssetPrefix + Hash(sourceData);
        if (sourceAssetId != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX image operation {operation.OperationId} source bytes no longer match the expected asset.", operation.SlidePartPath);
        var replacement = requestedAssets.Get(operation.ImageReplacement.AssetId);
        if (!sourcePart.ContentType.Equals(replacement.ContentType, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("unsupported_presentation_image", $"PPTX image operation {operation.OperationId} replacement must retain content type {sourcePart.ContentType}.", operation.SlidePartPath);

        var sourceImagePartPath = PartPath(sourcePart);
        var extension = Path.GetExtension(sourceImagePartPath);
        if (extension.Length is < 2 or > 12 || extension.Any(character => !char.IsAsciiLetterOrDigit(character) && character != '.'))
            throw new CodecException("unsupported_presentation_image", $"PPTX image operation {operation.OperationId} source part has an unsafe extension.", sourceImagePartPath);
        var replacementPath = $"ppt/media/office-kit-{replacement.Sha256[..24].ToLowerInvariant()}{extension.ToLowerInvariant()}";
        var sourceParts = PackageParts(sourceBytes);
        var addsPart = !sourceParts.TryGetValue(replacementPath, out var existing);
        if (!addsPart && !existing!.AsSpan().SequenceEqual(replacement.Data.Span))
            throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX replacement part path {replacementPath} already contains different bytes.", replacementPath);

        var usedRelationshipIds = slidePart.Parts.Select(pair => pair.RelationshipId)
            .Concat(slidePart.ExternalRelationships.Select(item => item.Id))
            .Concat(slidePart.HyperlinkRelationships.Select(item => item.Id))
            .Concat(slidePart.DataPartReferenceRelationships.Select(item => item.Id))
            .ToHashSet(StringComparer.Ordinal);
        var stem = $"rIdOfficeKitImage{replacement.Sha256[..16].ToLowerInvariant()}_";
        var replacementRelationshipId = Enumerable.Range(1, 1_000_000)
            .Select(index => stem + index)
            .FirstOrDefault(candidate => !usedRelationshipIds.Contains(candidate)) ??
            throw new CodecException("presentation_relationship_budget_exceeded", "PPTX image relationship ID allocation exceeded its bounded search.", operation.SlidePartPath);
        var relationshipPartPath = RelationshipPartPathFor(operation.SlidePartPath);
        if (!sourceParts.TryGetValue(relationshipPartPath, out var relationshipBytes))
            throw new CodecException("presentation_edit_target_missing", $"PPTX image operation {operation.OperationId} slide relationship part is missing.", relationshipPartPath);
        var relationshipType = SourceRelationshipType(relationshipBytes, relationshipId, relationshipPartPath);
        var replacementTarget = RelativePartTarget(operation.SlidePartPath, replacementPath);
        return new PptxImageEditPlanProof(
            replacement.Clone(),
            relationshipId,
            sourceImagePartPath,
            relationshipPartPath,
            relationshipType,
            replacementRelationshipId,
            replacementPath,
            replacementTarget,
            addsPart);
    }

    private static PptxImageEditPlanProof ProveImageSvgReplacement(
        byte[] sourceBytes,
        SlidePart slidePart,
        OpenXmlElement element,
        PresentationElement projectedElement,
        PresentationEditOperation operation,
        PptxAssetCatalog requestedAssets)
    {
        if (element is not P.Picture picture || projectedElement.ContentCase != PresentationElement.ContentOneofCase.Image ||
            projectedElement.Source.Editable != true || projectedElement.Image.AssetId != operation.ImageReplacement.AssetId ||
            projectedElement.Image.SvgAssetId != operation.ExpectedValue)
            throw new CodecException("unsupported_presentation_edit", $"PPTX SVG image operation {operation.OperationId} target is not a bounded editable picture with a source fallback.", operation.SlidePartPath);
        var relationshipId = SvgFallbackRelationshipId(picture);
        if (string.IsNullOrWhiteSpace(relationshipId))
            throw new CodecException("presentation_edit_target_missing", $"PPTX SVG image operation {operation.OperationId} source fallback relationship is missing.", operation.SlidePartPath);
        ImagePart sourcePart;
        try
        {
            sourcePart = slidePart.GetPartById(relationshipId) as ImagePart ??
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX SVG image operation {operation.OperationId} relationship is not an ImagePart.", operation.SlidePartPath);
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw new CodecException("presentation_edit_target_missing", $"PPTX SVG image operation {operation.OperationId} source relationship is missing.", operation.SlidePartPath, exception);
        }
        if (!sourcePart.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
            throw new CodecException("unsupported_presentation_image", $"PPTX SVG image operation {operation.OperationId} source fallback is not SVG.", operation.SlidePartPath);
        var sourceData = ReadOpenXmlPart(sourcePart);
        var sourceAssetId = PictureAssetPrefix + Hash(sourceData);
        if (sourceAssetId != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX SVG image operation {operation.OperationId} source bytes no longer match the expected fallback asset.", operation.SlidePartPath);
        var replacement = requestedAssets.Get(operation.ImageReplacement.SvgAssetId);
        if (!replacement.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
            throw new CodecException("unsupported_presentation_image", $"PPTX SVG image operation {operation.OperationId} replacement must use content type image/svg+xml.", operation.SlidePartPath);
        var sourceImagePartPath = PartPath(sourcePart);
        var extension = Path.GetExtension(sourceImagePartPath);
        if (extension.Length is < 2 or > 12 || extension.Any(character => !char.IsAsciiLetterOrDigit(character) && character != '.'))
            throw new CodecException("unsupported_presentation_image", $"PPTX SVG image operation {operation.OperationId} source part has an unsafe extension.", sourceImagePartPath);
        var replacementPath = $"ppt/media/office-kit-{replacement.Sha256[..24].ToLowerInvariant()}{extension.ToLowerInvariant()}";
        var sourceParts = PackageParts(sourceBytes);
        var addsPart = !sourceParts.TryGetValue(replacementPath, out var existing);
        if (!addsPart && !existing!.AsSpan().SequenceEqual(replacement.Data.Span))
            throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX replacement part path {replacementPath} already contains different bytes.", replacementPath);
        var usedRelationshipIds = slidePart.Parts.Select(pair => pair.RelationshipId)
            .Concat(slidePart.ExternalRelationships.Select(item => item.Id))
            .Concat(slidePart.HyperlinkRelationships.Select(item => item.Id))
            .Concat(slidePart.DataPartReferenceRelationships.Select(item => item.Id))
            .ToHashSet(StringComparer.Ordinal);
        var stem = $"rIdOfficeKitSvgImage{replacement.Sha256[..16].ToLowerInvariant()}_";
        var replacementRelationshipId = Enumerable.Range(1, 1_000_000)
            .Select(index => stem + index)
            .FirstOrDefault(candidate => !usedRelationshipIds.Contains(candidate)) ??
            throw new CodecException("presentation_relationship_budget_exceeded", "PPTX SVG relationship ID allocation exceeded its bounded search.", operation.SlidePartPath);
        var relationshipPartPath = RelationshipPartPathFor(operation.SlidePartPath);
        if (!sourceParts.TryGetValue(relationshipPartPath, out var relationshipBytes))
            throw new CodecException("presentation_edit_target_missing", $"PPTX SVG image operation {operation.OperationId} slide relationship part is missing.", relationshipPartPath);
        var relationshipType = SourceRelationshipType(relationshipBytes, relationshipId, relationshipPartPath);
        var replacementTarget = RelativePartTarget(operation.SlidePartPath, replacementPath);
        return new PptxImageEditPlanProof(
            replacement.Clone(),
            relationshipId,
            sourceImagePartPath,
            relationshipPartPath,
            relationshipType,
            replacementRelationshipId,
            replacementPath,
            replacementTarget,
            addsPart);
    }

    private static PptxXmlPatch CompileImageXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var image = proof.Image ?? throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} lost its package proof.", operation.SlidePartPath);
        if (elementRange.LocalName != "pic")
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} raw target is not p:pic.", operation.SlidePartPath);
        var namespaceByPrefix = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var declaration in NamespacePattern().Matches(xml).Cast<Match>())
        {
            var prefix = declaration.Groups["prefix"].Value;
            var uri = declaration.Groups["uri"].Value;
            if (namespaceByPrefix.TryGetValue(prefix, out var existingUri) && existingUri != uri)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} has an ambiguously rebound XML namespace prefix.", operation.SlidePartPath);
            namespaceByPrefix[prefix] = uri;
        }
        var elementXml = xml[elementRange.Start..elementRange.End];
        var tokens = XmlTokenPattern().Matches(elementXml).Cast<Match>().ToArray();
        var targetLocalName = LeafKind(operation) == "imageSvgAsset" ? "svgBlip" : "blip";
        var targetNamespace = LeafKind(operation) == "imageSvgAsset" ? SvgBlipNamespace : DrawingNamespace;
        var blips = tokens.Where(token => !token.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(token.Value) == targetLocalName &&
            namespaceByPrefix.TryGetValue(QualifiedPrefix(token.Value), out var uri) && uri == targetNamespace).ToArray();
        if (blips.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} requires one DrawingML blip token.", operation.SlidePartPath);
        var blip = blips[0];
        var embed = XmlAttributePattern().Matches(blip.Value).Cast<Match>().Where(attribute =>
        {
            var name = attribute.Groups["name"].Value;
            var prefix = QualifiedAttributePrefix(name);
            return LocalAttributeName(name) == "embed" && prefix.Length > 0 && namespaceByPrefix.TryGetValue(prefix, out var uri) &&
                uri is OfficeRelationshipsNamespace or StrictOfficeRelationshipsNamespace;
        }).ToArray();
        if (embed.Length != 1 || embed[0].Groups["value"].Value != image.SourceRelationshipId)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX image operation {operation.OperationId} embedded relationship no longer matches the source.", operation.SlidePartPath);

        if (LeafKind(operation) == "imageSvgAsset")
        {
            var start = blip.Index + embed[0].Groups["value"].Index;
            var end = start + embed[0].Groups["value"].Length;
            var replacementXml = elementXml[..start] + image.ReplacementRelationshipId + elementXml[end..];
            return new PptxXmlPatch(operation, elementRange.Start, elementRange.End, replacementXml, proof.SourceElementSha256, proof.MutationPartPath);
        }

        var localPatches = new List<(int Start, int End, string Replacement)>
        {
            (blip.Index + embed[0].Groups["value"].Index, blip.Index + embed[0].Groups["value"].Index + embed[0].Groups["value"].Length, image.ReplacementRelationshipId),
        };
        var crops = tokens.Where(token => !token.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(token.Value) == "srcRect" &&
            namespaceByPrefix.TryGetValue(QualifiedPrefix(token.Value), out var uri) && uri == DrawingNamespace).ToArray();
        if (crops.Length > 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} has ambiguous source rectangles.", operation.SlidePartPath);
        if (operation.ImageReplacement.Crop is { } crop)
        {
            if (crops.Length == 1)
            {
                var replacement = UpdateCropToken(crops[0].Value, crop, operation);
                localPatches.Add((crops[0].Index, crops[0].Index + crops[0].Length, replacement));
            }
            else
            {
                var prefix = QualifiedPrefix(blip.Value);
                var replacement = BuildCropToken(prefix, crop);
                var insertion = XmlElementEnd(tokens, blip, operation);
                localPatches.Add((insertion, insertion, replacement));
            }
        }
        else if (crops.Length == 1)
        {
            localPatches.Add((crops[0].Index, crops[0].Index + crops[0].Length, string.Empty));
        }
        var replacementElement = elementXml;
        foreach (var patch in localPatches.OrderByDescending(item => item.Start))
            replacementElement = replacementElement[..patch.Start] + patch.Replacement + replacementElement[patch.End..];
        return new PptxXmlPatch(operation, elementRange.Start, elementRange.End, replacementElement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static void ApplyImagePackagePatches(
        IReadOnlyDictionary<string, byte[]> sourceParts,
        IReadOnlyList<PptxEditPlanProof> proofs,
        IDictionary<string, byte[]> patchedParts,
        IDictionary<string, byte[]> addedParts)
    {
        var imageProofs = proofs.Where(proof => proof.Image is not null).ToArray();
        foreach (var group in imageProofs.GroupBy(proof => proof.Image!.RelationshipPartPath, StringComparer.OrdinalIgnoreCase))
        {
            var path = group.Key;
            if (patchedParts.ContainsKey(path))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX image operations overlap another mutation of {path}.", path);
            patchedParts.Add(path, AppendImageRelationships(sourceParts[path], group.Select(proof => proof.Image!).ToArray(), path));
        }
        foreach (var image in imageProofs.Select(proof => proof.Image!).Where(image => image.AddsPart))
        {
            var bytes = image.Replacement.Data.ToByteArray();
            if (addedParts.TryGetValue(image.ReplacementPartPath, out var existing))
            {
                if (!existing.AsSpan().SequenceEqual(bytes))
                    throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX image operations disagree on added part {image.ReplacementPartPath}.", image.ReplacementPartPath);
                continue;
            }
            addedParts.Add(image.ReplacementPartPath, bytes);
        }
    }

    private static void VerifyImageReplacement(
        SlidePart slidePart,
        OpenXmlElement element,
        PresentationEditOperation operation,
        PresentationEditOperationResult result)
    {
        if (element is not P.Picture picture)
            throw new CodecException("presentation_edit_verification_failed", $"PPTX image operation {operation.OperationId} target is no longer a picture.", operation.SlidePartPath);
        var relationshipId = LeafKind(operation) == "imageSvgAsset"
            ? SvgFallbackRelationshipId(picture)
            : picture.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(relationshipId))
            throw new CodecException("presentation_edit_verification_failed", $"PPTX image operation {operation.OperationId} output relationship is missing.", operation.SlidePartPath);
        ImagePart imagePart;
        try
        {
            imagePart = slidePart.GetPartById(relationshipId!) as ImagePart ??
                throw new CodecException("presentation_edit_verification_failed", $"PPTX image operation {operation.OperationId} output relationship is not an ImagePart.", operation.SlidePartPath);
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw new CodecException("presentation_edit_verification_failed", $"PPTX image operation {operation.OperationId} output relationship is missing.", operation.SlidePartPath, exception);
        }
        if (PictureAssetPrefix + Hash(ReadOpenXmlPart(imagePart)) != operation.Value ||
            LeafKind(operation) == "imageAsset" && !SameCrop(picture.BlipFill?.GetFirstChild<A.SourceRectangle>(), operation.ImageReplacement.Crop))
            throw new CodecException("presentation_edit_verification_failed", $"PPTX image operation {operation.OperationId} did not retain its asset and crop.", operation.SlidePartPath);
        result.OutputElementSha256 = HashElement(element);
    }

    private static OpenXmlElement? SvgFallbackElement(P.Picture picture) => picture.BlipFill?.GetFirstChild<A.Blip>()?.ChildElements
        .Where(child => child.LocalName == "extLst" && child.NamespaceUri == DrawingNamespace)
        .SelectMany(child => child.ChildElements)
        .Where(child => child.LocalName == "ext" && child.NamespaceUri == DrawingNamespace &&
            child.GetAttributes().Any(attribute => attribute.LocalName == "uri" && attribute.NamespaceUri.Length == 0 && attribute.Value == SvgExtensionUri))
        .SelectMany(child => child.ChildElements)
        .SingleOrDefault(child => child.LocalName == "svgBlip" && child.NamespaceUri == SvgBlipNamespace);

    private static string? SvgFallbackRelationshipId(P.Picture picture)
    {
        var fallback = SvgFallbackElement(picture);
        if (fallback is null) return null;
        var embeds = fallback.GetAttributes()
            .Where(attribute => attribute.LocalName == "embed" &&
                attribute.NamespaceUri is OfficeRelationshipsNamespace or StrictOfficeRelationshipsNamespace)
            .ToArray();
        return embeds.Length == 1 ? embeds[0].Value : string.Empty;
    }

    private static byte[] AppendImageRelationships(byte[] sourcePart, IReadOnlyList<PptxImageEditPlanProof> images, string partPath)
    {
        var (xml, bomBytes) = DecodeXml(sourcePart);
        var tokens = XmlTokenPattern().Matches(xml).Cast<Match>().ToArray();
        var closing = tokens.Where(token => token.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(token.Value) == "Relationships").ToArray();
        if (closing.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", "PPTX relationship part has no unique Relationships root.", partPath);
        var existingIds = tokens.Where(token => !token.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(token.Value) == "Relationship")
            .SelectMany(token => XmlAttributePattern().Matches(token.Value).Cast<Match>())
            .Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "Id")
            .Select(attribute => attribute.Groups["value"].Value)
            .ToHashSet(StringComparer.Ordinal);
        var additions = new StringBuilder();
        foreach (var image in images.OrderBy(item => item.ReplacementRelationshipId, StringComparer.Ordinal))
        {
            if (!existingIds.Add(image.ReplacementRelationshipId))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX relationship ID {image.ReplacementRelationshipId} already exists.", partPath);
            var sourceTag = RelationshipToken(tokens, image.SourceRelationshipId, partPath);
            var qualifiedName = QualifiedName(sourceTag.Value);
            additions.Append('<').Append(qualifiedName)
                .Append(" Id=\"").Append(EscapeAttribute(image.ReplacementRelationshipId)).Append('"')
                .Append(" Type=\"").Append(EscapeAttribute(image.RelationshipType)).Append('"')
                .Append(" Target=\"").Append(EscapeAttribute(image.ReplacementTarget)).Append("\"/>");
        }
        var output = xml.Insert(closing[0].Index, additions.ToString());
        var encoded = StrictUtf8.GetBytes(output);
        return bomBytes == 0 ? encoded : StrictUtf8.GetPreamble().Concat(encoded).ToArray();
    }

    private static Match RelationshipToken(IReadOnlyList<Match> tokens, string relationshipId, string partPath)
    {
        var matches = tokens.Where(token => !token.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(token.Value) == "Relationship")
            .Where(token => XmlAttributePattern().Matches(token.Value).Cast<Match>().Any(attribute =>
                LocalAttributeName(attribute.Groups["name"].Value) == "Id" && attribute.Groups["value"].Value == relationshipId))
            .ToArray();
        return matches.Length == 1 ? matches[0] :
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX source relationship {relationshipId} is missing or ambiguous.", partPath);
    }

    private static string SourceRelationshipType(byte[] relationshipBytes, string relationshipId, string partPath)
    {
        var (xml, _) = DecodeXml(relationshipBytes);
        var token = RelationshipToken(XmlTokenPattern().Matches(xml).Cast<Match>().ToArray(), relationshipId, partPath);
        var types = XmlAttributePattern().Matches(token.Value).Cast<Match>()
            .Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "Type")
            .Select(attribute => attribute.Groups["value"].Value)
            .ToArray();
        if (types.Length != 1 || !(types[0].EndsWith("/image", StringComparison.Ordinal) || types[0].EndsWith("/image", StringComparison.OrdinalIgnoreCase)))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX source relationship {relationshipId} is not an image relationship.", partPath);
        return types[0];
    }

    private static bool ValidImageCrop(PresentationImageCrop crop) =>
        crop.LeftThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.TopThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.RightThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.BottomThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.LeftThousandthPercent + crop.RightThousandthPercent < 100_000 &&
        crop.TopThousandthPercent + crop.BottomThousandthPercent < 100_000;

    private static bool SameCrop(A.SourceRectangle? actual, PresentationImageCrop? expected) => expected is null
        ? actual is null
        : actual is not null &&
          (actual.Left?.Value ?? 0) == expected.LeftThousandthPercent &&
          (actual.Top?.Value ?? 0) == expected.TopThousandthPercent &&
          (actual.Right?.Value ?? 0) == expected.RightThousandthPercent &&
          (actual.Bottom?.Value ?? 0) == expected.BottomThousandthPercent;

    private static string BuildCropToken(string prefix, PresentationImageCrop crop)
    {
        var name = prefix.Length == 0 ? "srcRect" : $"{prefix}:srcRect";
        var declaration = prefix.Length == 0
            ? $" xmlns=\"{DrawingNamespace}\""
            : $" xmlns:{prefix}=\"{DrawingNamespace}\"";
        return $"<{name}{declaration} l=\"{crop.LeftThousandthPercent}\" t=\"{crop.TopThousandthPercent}\" r=\"{crop.RightThousandthPercent}\" b=\"{crop.BottomThousandthPercent}\"/>";
    }

    private static string UpdateCropToken(string token, PresentationImageCrop crop, PresentationEditOperation operation)
    {
        if (!token.TrimEnd().EndsWith("/>", StringComparison.Ordinal))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} source rectangle is not an empty element.", operation.SlidePartPath);
        var values = new Dictionary<string, long>(StringComparer.Ordinal)
        {
            ["l"] = crop.LeftThousandthPercent,
            ["t"] = crop.TopThousandthPercent,
            ["r"] = crop.RightThousandthPercent,
            ["b"] = crop.BottomThousandthPercent,
        };
        var patches = new List<(int Start, int End, string Replacement)>();
        var present = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match attribute in XmlAttributePattern().Matches(token))
        {
            var name = attribute.Groups["name"].Value;
            if (!values.TryGetValue(name, out var value)) continue;
            if (!present.Add(name))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} source rectangle contains duplicate {name} attributes.", operation.SlidePartPath);
            var group = attribute.Groups["value"];
            patches.Add((group.Index, group.Index + group.Length, value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        var output = token;
        foreach (var patch in patches.OrderByDescending(item => item.Start))
            output = output[..patch.Start] + patch.Replacement + output[patch.End..];
        var missing = values.Where(pair => !present.Contains(pair.Key)).Select(pair => $" {pair.Key}=\"{pair.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}\"");
        return output.Insert(output.LastIndexOf("/>", StringComparison.Ordinal), string.Concat(missing));
    }

    private static int XmlElementEnd(
        IReadOnlyList<Match> tokens,
        Match opening,
        PresentationEditOperation operation)
    {
        if (opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal)) return opening.Index + opening.Length;
        var qualifiedName = QualifiedName(opening.Value);
        var depth = 1;
        foreach (var token in tokens.Where(token => token.Index >= opening.Index + opening.Length))
        {
            if (QualifiedName(token.Value) != qualifiedName) continue;
            if (token.Value.StartsWith("</", StringComparison.Ordinal)) depth--;
            else if (!token.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal)) depth++;
            if (depth == 0) return token.Index + token.Length;
        }
        throw new CodecException("presentation_edit_target_mismatch", $"PPTX image operation {operation.OperationId} DrawingML blip is unbalanced.", operation.SlidePartPath);
    }

    private static byte[] ReadOpenXmlPart(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static string RelationshipPartPathFor(string ownerPath)
    {
        var separator = ownerPath.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : ownerPath[..separator];
        var fileName = separator < 0 ? ownerPath : ownerPath[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }

    private static string RelativePartTarget(string ownerPath, string targetPath)
    {
        var separator = ownerPath.LastIndexOf('/');
        var ownerDirectory = separator < 0 ? string.Empty : ownerPath[..(separator + 1)];
        var baseUri = new Uri("https://officekit.invalid/" + ownerDirectory, UriKind.Absolute);
        var targetUri = new Uri("https://officekit.invalid/" + targetPath, UriKind.Absolute);
        return Uri.UnescapeDataString(baseUri.MakeRelativeUri(targetUri).OriginalString);
    }

    private static string QualifiedName(string tag)
    {
        var match = Regex.Match(tag, "^</?(?<name>(?:[A-Za-z_][\\w.-]*:)?[A-Za-z_][\\w.-]*)\\b", RegexOptions.CultureInvariant);
        return match.Success ? match.Groups["name"].Value : string.Empty;
    }

    private static string QualifiedPrefix(string tag)
    {
        var name = QualifiedName(tag);
        var separator = name.IndexOf(':');
        return separator < 0 ? string.Empty : name[..separator];
    }

    private static string QualifiedAttributePrefix(string name)
    {
        var separator = name.IndexOf(':');
        return separator < 0 ? string.Empty : name[..separator];
    }

    private static string EscapeAttribute(string value) => new System.Xml.Linq.XAttribute("v", value).ToString()[3..^1];
}
