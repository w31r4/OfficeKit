using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxEditPlanOutput(
    byte[] File,
    PresentationEditPlanResult Result,
    IReadOnlyList<Diagnostic> Diagnostics);

internal sealed record PptxEditPlanProof(
    PresentationEditOperation Operation,
    string SourceElementSha256,
    string MutationPartPath,
    uint? RawTextOrdinal = null,
    PptxImageEditPlanProof? Image = null,
    PptxElementDeletionPlan? Deletion = null);

internal sealed record PptxXmlPatch(
    PresentationEditOperation Operation,
    int Start,
    int End,
    string Replacement,
    string SourceElementSha256,
    string MutationPartPath);

// Applies a finite, source-bound edit plan directly to the original XML token
// stream. The Open XML SDK is used only as an independent structural oracle;
// it never serializes the mutated SlidePart.
internal static partial class PptxEditPlanCodec
{
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    [GeneratedRegex("^ppt/slides/slide[1-9][0-9]*[.]xml$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex SlidePathPattern();

    [GeneratedRegex("^ppt/(?:[A-Za-z0-9_.-]+/)*[A-Za-z0-9_.-]+[.]xml$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex DependentXmlPartPathPattern();

    [GeneratedRegex("^ppt/(?:[A-Za-z0-9_.-]+/)*[A-Za-z0-9_.-]+[.]xlsx$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex EmbeddedPackagePartPathPattern();

    [GeneratedRegex("^xl/worksheets/[A-Za-z0-9_.-]+[.]xml$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex EmbeddedWorksheetPartPathPattern();

    [GeneratedRegex("^[A-Z]{1,3}[1-9][0-9]*$", RegexOptions.CultureInvariant)]
    private static partial Regex CellReferencePattern();

    [GeneratedRegex("<!--.*?-->|<!\\[CDATA\\[.*?\\]\\]>|<\\?.*?\\?>|</?(?:[A-Za-z_][\\w.-]*:)?[A-Za-z_][\\w.-]*\\b[^>]*>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex XmlTokenPattern();

    [GeneratedRegex("^</?(?:[A-Za-z_][\\w.-]*:)?(?<name>[A-Za-z_][\\w.-]*)\\b", RegexOptions.CultureInvariant)]
    private static partial Regex XmlLocalNamePattern();

    [GeneratedRegex("xmlns:(?<prefix>[A-Za-z_][\\w.-]*)\\s*=\\s*(?<quote>['\"])(?<uri>.*?)\\k<quote>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex NamespacePattern();

    [GeneratedRegex("(?:(?<open><(?<prefix>[A-Za-z_][\\w.-]*):t\\b[^>]*?\\/\\s*>)|(?<open><(?<prefix>[A-Za-z_][\\w.-]*):t\\b[^>]*>)(?<value>.*?)(?<close></\\k<prefix>:t\\s*>))", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex TextLeafPattern();

    [GeneratedRegex("<(?<prefix>[A-Za-z_][\\w.-]*):tc\\b[^>]*>.*?</\\k<prefix>:tc\\s*>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex TableCellPattern();

    [GeneratedRegex("<(?<prefix>[A-Za-z_][\\w.-]*):r\\b[^>]*>.*?</\\k<prefix>:r\\s*>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex TextRunPattern();

    [GeneratedRegex("<(?<prefix>[A-Za-z_][\\w.-]*):sp\\b[^>]*>.*?</\\k<prefix>:sp\\s*>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex ShapeTextPattern();

    [GeneratedRegex("(?<name>(?:[A-Za-z_][\\w.-]*:)?[A-Za-z_][\\w.-]*)\\s*=\\s*(?<quote>['\"])(?<value>.*?)\\k<quote>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex XmlAttributePattern();

    [GeneratedRegex("\\bxml:space\\s*=\\s*(?<quote>['\"])preserve\\k<quote>", RegexOptions.CultureInvariant)]
    private static partial Regex PreserveSpacePattern();

    internal static PptxEditPlanOutput Apply(
        byte[] sourceBytes,
        PresentationEditPlanRequest request,
        EffectiveCodecLimits limits)
    {
        ValidateRequest(sourceBytes, request, limits);
        var sourceHash = Hash(sourceBytes);
        _ = PackageGuards.ValidateAndCollectOpaque(sourceBytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        var sourceProjection = PptxCodec.Import(sourceBytes, limits).Artifact.Presentation;
        var proofs = ProveOperations(sourceBytes, request, sourceProjection, limits);
        var sourceParts = PackageParts(sourceBytes);
        var patchedParts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var addedParts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var removedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var results = new List<PresentationEditOperationResult>();

        foreach (var group in proofs.GroupBy(proof => proof.MutationPartPath, StringComparer.OrdinalIgnoreCase))
        {
            var partPath = group.Key;
            var sourcePart = sourceParts[partPath];
            var patches = CompileXmlPatches(sourcePart, group.ToArray());
            var outputPart = ApplyPatches(sourcePart, patches, results);
            if (sourcePart.AsSpan().SequenceEqual(outputPart))
                throw new CodecException("presentation_edit_plan_noop", $"PPTX edit plan produced no change for {partPath}.", partPath);
            patchedParts.Add(partPath, outputPart);
        }
        ApplyEmbeddedWorkbookPatches(sourceParts, proofs, patchedParts, results);
        ApplyImagePackagePatches(sourceParts, proofs, patchedParts, addedParts);
        ApplyElementDeletionPackagePatches(sourceParts, proofs, patchedParts, removedParts);

        var outputBytes = RewriteParts(sourceBytes, patchedParts, addedParts, removedParts);
        var outputParts = PackageParts(outputBytes);
        var changedParts = ChangedParts(sourceParts, outputParts);
        var expectedParts = patchedParts.Keys.Concat(addedParts.Keys).Concat(removedParts).OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
        if (!changedParts.SequenceEqual(expectedParts, StringComparer.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_edit_plan_scope_violation",
                $"PPTX edit plan changed unexpected OPC parts: {string.Join(", ", changedParts)}.");

        foreach (var (path, expected) in patchedParts)
        {
            var actual = RequiredPart(outputParts, path);
            if (!actual.AsSpan().SequenceEqual(expected))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX edit plan output for {path} differs from the compiled token patch.", path);
        }
        foreach (var (path, expected) in addedParts)
        {
            var actual = RequiredPart(outputParts, path);
            if (!actual.AsSpan().SequenceEqual(expected))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX edit plan added part {path} with unexpected bytes.", path);
        }
        foreach (var path in removedParts)
            if (outputParts.ContainsKey(path))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX edit plan failed to remove part {path}.", path);

        _ = PackageGuards.ValidateAndCollectOpaque(outputBytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        var sourceValidationWarnings = PptxCodec.ValidateEditPlanOutput(sourceBytes, outputBytes, limits);
        VerifyOutput(outputBytes, request, results, limits);

        var result = new PresentationEditPlanResult
        {
            SourceSha256 = sourceHash,
            OutputSha256 = Hash(outputBytes),
        };
        result.ChangedParts.Add(changedParts);
        result.Operations.Add(results.OrderBy(item => item.OperationId, StringComparer.Ordinal));
        var diagnostics = new List<Diagnostic>();
        if (sourceValidationWarnings > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "source_openxml_validation_warnings_preserved",
                $"Preserved {sourceValidationWarnings} pre-existing Office 2021 validation warning(s) from the source package; the edit plan introduced none."));
        return new PptxEditPlanOutput(outputBytes, result, diagnostics);
    }

    private static void ValidateRequest(byte[] sourceBytes, PresentationEditPlanRequest request, EffectiveCodecLimits limits)
    {
        if (request is null) throw new CodecException("missing_presentation_edit_plan", "PPTX edit plan is required.");
        var sourceHash = Hash(sourceBytes);
        if (!IsSha256(request.ExpectedSourceSha256) || !sourceHash.Equals(request.ExpectedSourceSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("presentation_source_hash_mismatch", "PPTX edit plan requires expected_source_sha256 to match the exact input bytes.");
        if (request.Operations.Count == 0)
            throw new CodecException("empty_presentation_edit_plan", "PPTX edit plan requires at least one operation.");
        if ((ulong)request.Operations.Count > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"PPTX edit plan has {request.Operations.Count} operations and exceeds max_cells ({limits.MaxCells}).");
        if (request.Operations.Select(operation => operation.OperationId).Distinct(StringComparer.Ordinal).Count() != request.Operations.Count)
            throw new CodecException("duplicate_presentation_edit_operation", "PPTX edit plan operation IDs must be unique.");
        if (request.Operations.GroupBy(operation => (MutationPartPath(operation).ToLowerInvariant(), ShapeTreePathKey(operation), LeafKind(operation), LeafIndexKey(operation))).Any(group => group.Count() > 1))
            throw new CodecException("duplicate_presentation_edit_target", "PPTX edit plan cannot edit the same native leaf twice.");
        ulong textBudget = 0;
        foreach (var operation in request.Operations)
        {
            if (string.IsNullOrWhiteSpace(operation.OperationId) || operation.OperationId.Length > 160 ||
                string.IsNullOrWhiteSpace(operation.SlideId) || operation.SlideId.Length > 512 ||
                string.IsNullOrWhiteSpace(operation.TargetId) || operation.TargetId.Length > 512)
                throw new CodecException("invalid_presentation_edit_operation", "PPTX edit operation IDs must be non-empty and bounded.");
            if (!SlidePathPattern().IsMatch(operation.SlidePartPath) || operation.SlidePartPath.Contains("..", StringComparison.Ordinal))
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid slide part path.");
            var shapeTreePath = ShapeTreePath(operation);
            if (shapeTreePath.Count > 32 || shapeTreePath[0] != operation.ShapeTreeIndex)
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid shape-tree path.");
            var leafKind = LeafKind(operation);
            if (leafKind is not ("text" or "tableCellText" or "nativeText" or "fillRgb" or "lineRgb" or "lineScheme" or "lineWidthEmu" or "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "imageAsset" or "imageSvgAsset" or "chartTitleText" or "chartDataValue" or "diagramText" or "deleteElement"))
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported leaf kind {leafKind}.");
            if (!IsSha256(operation.ExpectedSlideSha256) || !IsSha256(operation.ExpectedElementSha256) ||
                !IsSha256(operation.ExpectedSemanticSha256) || !IsSha256(operation.ExpectedTextSha256))
                throw new CodecException("invalid_presentation_edit_precondition", $"PPTX edit operation {operation.OperationId} requires SHA-256 preconditions.");
            if (!Hash(Encoding.UTF8.GetBytes(operation.ExpectedValue)).Equals(operation.ExpectedTextSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_text_hash_mismatch", $"PPTX edit operation {operation.OperationId} expected text does not match expected_text_sha256.");
            if (leafKind is "chartTitleText" or "chartDataValue" or "diagramText")
            {
                if (!DependentXmlPartPathPattern().IsMatch(operation.TargetPartPath) || operation.TargetPartPath.Contains("..", StringComparison.Ordinal) ||
                    !IsSha256(operation.ExpectedTargetPartSha256) || string.IsNullOrWhiteSpace(operation.RelationshipId) || operation.RelationshipId.Length > 255)
                    throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid source-bound dependent-part binding.");
                if (leafKind == "chartDataValue") ValidateEmbeddedWorkbookBinding(operation);
                if (leafKind is "chartTitleText" or "diagramText" && HasEmbeddedWorkbookBinding(operation))
                    throw new CodecException("invalid_presentation_edit_target", $"PPTX {leafKind} operation {operation.OperationId} cannot attach an embedded-workbook data binding.");
            }
            if (leafKind == "diagramText")
            {
                if (string.IsNullOrWhiteSpace(operation.DiagramModelId) || operation.DiagramModelId.Length > 1_024 || operation.DiagramModelId.Any(char.IsControl))
                    throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid SmartArt model ID binding.");
            }
            else if (HasDiagramBinding(operation))
            {
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} cannot attach a SmartArt run binding to {leafKind}.");
            }
            if (leafKind is "imageAsset" or "imageSvgAsset")
            {
                if (operation.ImageReplacement is null ||
                    (leafKind == "imageAsset" && operation.ImageReplacement.AssetId != operation.Value) ||
                    (leafKind == "imageSvgAsset" && operation.ImageReplacement.SvgAssetId != operation.Value))
                    throw new CodecException("invalid_presentation_edit_target", $"PPTX image operation {operation.OperationId} requires one replacement asset matching value.");
                ValidateImageReplacement(operation);
            }
            else if (operation.ImageReplacement is not null)
            {
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} cannot attach an image replacement to {leafKind}.");
            }
            if (leafKind == "deleteElement")
            {
                if (operation.ElementDeletion is null || operation.ElementDeletion.ExpectedNativeId == 0 || shapeTreePath.Count != 1 ||
                    operation.ExpectedValue != operation.TargetId || operation.Value.Length != 0)
                    throw new CodecException("invalid_presentation_edit_target", $"PPTX element deletion {operation.OperationId} requires one top-level codec-issued native identity binding.");
            }
            else if (operation.ElementDeletion is not null)
            {
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} cannot attach an element deletion to {leafKind}.");
            }
            if (leafKind != "chartDataValue" && (operation.ChartSeriesIndex != 0 || operation.ChartPointIndex != 0))
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} cannot attach chart-data indices to {leafKind}.");
            if (leafKind == "chartDataValue" && (!ValidFiniteNumber(operation.ExpectedValue) || !ValidFiniteNumber(operation.Value)))
                throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} chart data value must be a finite numeric token.");
            if (leafKind is not ("chartTitleText" or "chartDataValue" or "diagramText") &&
                (!string.IsNullOrEmpty(operation.TargetPartPath) || !string.IsNullOrEmpty(operation.ExpectedTargetPartSha256) || !string.IsNullOrEmpty(operation.RelationshipId) || HasEmbeddedWorkbookBinding(operation)))
            {
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} cannot attach a dependent-part binding to {leafKind}.");
            }
            if (operation.ExpectedValue == operation.Value)
                throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its target value.");
            if (leafKind is "fillRgb" or "lineRgb")
            {
                var expected = PptxColor.Normalize(operation.ExpectedValue);
                var requested = PptxColor.Normalize(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its target color.");
            }
            if (leafKind == "lineScheme")
            {
                var expected = PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its target color.");
            }
            if (leafKind.EndsWith("Emu", StringComparison.Ordinal))
            {
                if (!long.TryParse(operation.ExpectedValue, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out _) ||
                    !long.TryParse(operation.Value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var requested) ||
                    (leafKind is "widthEmu" or "heightEmu" && requested <= 0))
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} has an invalid geometry scalar.");
                if (leafKind == "lineWidthEmu" && (requested < 0 || requested > 20_116_800))
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} has an invalid line width.");
            }
            textBudget = checked(textBudget + (ulong)operation.ExpectedValue.Length + (ulong)operation.Value.Length);
            if (textBudget > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"PPTX edit-plan text exceeds max_cells ({limits.MaxCells}).");
        }
        var requestedAssetIds = request.Operations
            .Where(operation => LeafKind(operation) is "imageAsset" or "imageSvgAsset")
            .Select(operation => LeafKind(operation) == "imageSvgAsset"
                ? operation.ImageReplacement.SvgAssetId
                : operation.ImageReplacement.AssetId)
            .ToHashSet(StringComparer.Ordinal);
        var suppliedAssetIds = request.Assets.Select(asset => asset.Id).ToHashSet(StringComparer.Ordinal);
        if (!requestedAssetIds.SetEquals(suppliedAssetIds) || request.Assets.Count != suppliedAssetIds.Count)
            throw new CodecException("invalid_presentation_asset", "PPTX edit plan assets must match the imageAsset operation references exactly.");
        var deletedRoots = request.Operations
            .Where(operation => LeafKind(operation) == "deleteElement")
            .Select(operation => (operation.SlidePartPath.ToLowerInvariant(), ShapeTreePath(operation)[0]))
            .ToHashSet();
        if (request.Operations.Any(operation => LeafKind(operation) != "deleteElement" &&
                deletedRoots.Contains((operation.SlidePartPath.ToLowerInvariant(), ShapeTreePath(operation)[0]))))
            throw new CodecException("invalid_presentation_edit_target", "PPTX edit plan cannot mutate a native element that it also deletes.");
    }

    private static IReadOnlyList<PptxEditPlanProof> ProveOperations(
        byte[] sourceBytes,
        PresentationEditPlanRequest request,
        PresentationArtifact sourceProjection,
        EffectiveCodecLimits limits)
    {
        using var stream = new MemoryStream(sourceBytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false, new OpenSettings { AutoSave = false });
        var presentationPart = package.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.");
        var slideByPath = presentationPart.SlideParts.ToDictionary(PartPath, StringComparer.OrdinalIgnoreCase);
        var projectedSlideByPath = sourceProjection.Slides.ToDictionary(slide => slide.Source.PartPath, StringComparer.OrdinalIgnoreCase);
        var proofs = new List<PptxEditPlanProof>();
        var requestedAssets = new PptxAssetCatalog(request.Assets, limits);
        foreach (var operation in request.Operations)
        {
            if (!slideByPath.TryGetValue(operation.SlidePartPath, out var slidePart))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} slide part was not found.", operation.SlidePartPath);
            if (!projectedSlideByPath.TryGetValue(operation.SlidePartPath, out var projectedSlide) || projectedSlide.Id != operation.SlideId)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} slide identity does not match the imported source projection.", operation.SlidePartPath);
            if (!HashElement(slidePart.Slide!).Equals(operation.ExpectedSlideSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_slide_hash_mismatch", $"PPTX edit operation {operation.OperationId} slide XML changed after planning.", operation.SlidePartPath);
            var tree = slidePart.Slide?.CommonSlideData?.ShapeTree ??
                throw new CodecException("missing_presentation_shape_tree", "PPTX edit target slide has no shape tree.", operation.SlidePartPath);
            var path = ShapeTreePath(operation);
            var projectedElement = ResolveProjectedElement(projectedSlide.Elements, path, operation);
            if (projectedElement.Id != operation.TargetId)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} semantic target binding does not match the imported source projection.", operation.SlidePartPath);
            if (!projectedElement.Source.ElementSha256.Equals(operation.ExpectedElementSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_element_binding_mismatch", $"PPTX edit operation {operation.OperationId} element hash does not match the imported source projection.", operation.SlidePartPath);
            if (!projectedElement.Source.SemanticSha256.Equals(operation.ExpectedSemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_semantic_binding_mismatch", $"PPTX edit operation {operation.OperationId} semantic hash does not match the imported source projection.", operation.SlidePartPath);
            var element = ResolveShapeTreeElement(tree, path, operation);
            var elementHash = HashElement(element);
            if (!elementHash.Equals(operation.ExpectedElementSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_element_binding_mismatch", $"PPTX edit operation {operation.OperationId} target element changed after planning.", operation.SlidePartPath);
            if (LeafKind(operation) == "diagramText")
            {
                if (element is not P.GraphicFrame || projectedElement.ContentCase != PresentationElement.ContentOneofCase.Opaque ||
                    projectedElement.Opaque.DiagramText is null ||
                    !PptxDiagramTextCodec.TryResolveForEditPlan(element, slidePart, out var diagram) ||
                    !PptxNativeObjectCatalog.HasUniqueInboundRelationship(presentationPart, diagram.Part) ||
                    !PptxDiagramTextCodec.SameEditBinding(projectedElement.Opaque.DiagramText, diagram.Binding) ||
                    !operation.TargetPartPath.Equals(diagram.Binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
                    !operation.ExpectedTargetPartSha256.Equals(diagram.Binding.SourceSha256, StringComparison.OrdinalIgnoreCase) ||
                    operation.RelationshipId != diagram.Binding.RelationshipId)
                    throw new CodecException("presentation_diagram_text_binding_mismatch", $"PPTX edit operation {operation.OperationId} no longer resolves to its unique source-bound DiagramDataPart.", operation.SlidePartPath);
                if (operation.TextLeafIndex >= (uint)diagram.Leaves.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} SmartArt text leaf index is out of range.", diagram.Binding.PartPath);
                var leaf = diagram.Leaves[(int)operation.TextLeafIndex];
                if (leaf.ModelId != operation.DiagramModelId || leaf.RunIndex != operation.DiagramRunIndex || leaf.Text != operation.ExpectedValue)
                    throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} SmartArt run no longer matches its node/run binding.", diagram.Binding.PartPath);
                proofs.Add(new PptxEditPlanProof(operation, elementHash, diagram.Binding.PartPath, leaf.RawTextOrdinal));
                continue;
            }
            if (LeafKind(operation) is "chartTitleText" or "chartDataValue")
            {
                if (element is not P.GraphicFrame || projectedElement.ContentCase != PresentationElement.ContentOneofCase.Opaque ||
                    projectedElement.Opaque.NativeChart is null ||
                    !PptxNativeChartLeafCodec.TryResolve(element, slidePart, limits, out var chart) ||
                    !PptxNativeObjectCatalog.HasUniqueInboundRelationship(presentationPart, chart.Part) ||
                    !operation.TargetPartPath.Equals(chart.Binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
                    !operation.ExpectedTargetPartSha256.Equals(chart.Binding.SourceSha256, StringComparison.OrdinalIgnoreCase) ||
                    operation.RelationshipId != chart.Binding.RelationshipId)
                    throw new CodecException("presentation_chart_binding_mismatch", $"PPTX edit operation {operation.OperationId} no longer resolves to its unique source-bound ChartPart.", operation.SlidePartPath);
                if (LeafKind(operation) == "chartDataValue")
                {
                    if (!PptxNativeChartLeafCodec.SameBinding(projectedElement.Opaque.NativeChart, chart.Binding) || chart.Data is null ||
                        !PptxNativeObjectCatalog.HasUniqueInboundRelationship(presentationPart, chart.Data.Part) ||
                        !operation.EmbeddedPackagePartPath.Equals(chart.Binding.EmbeddedPackagePartPath, StringComparison.OrdinalIgnoreCase) ||
                        !operation.ExpectedEmbeddedPackageSha256.Equals(chart.Binding.EmbeddedPackageSourceSha256, StringComparison.OrdinalIgnoreCase) ||
                        operation.EmbeddedPackageRelationshipId != chart.Binding.EmbeddedPackageRelationshipId)
                        throw new CodecException("presentation_chart_data_binding_mismatch", $"PPTX edit operation {operation.OperationId} no longer resolves to its unique source-bound embedded workbook.", operation.TargetPartPath);
                    var point = chart.Data.Points.SingleOrDefault(candidate =>
                        candidate.Binding.SeriesIndex == operation.ChartSeriesIndex && candidate.Binding.PointIndex == operation.ChartPointIndex);
                    if (point is null || point.Binding.Value != operation.ExpectedValue || point.Binding.Formula != operation.ChartFormula ||
                        !point.Binding.WorksheetPartPath.Equals(operation.EmbeddedWorksheetPartPath, StringComparison.OrdinalIgnoreCase) ||
                        !point.Binding.WorksheetSourceSha256.Equals(operation.ExpectedEmbeddedWorksheetSha256, StringComparison.OrdinalIgnoreCase) ||
                        !point.Binding.CellReference.Equals(operation.EmbeddedCellReference, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} chart data point no longer matches its cache/workbook binding.", operation.TargetPartPath);
                    proofs.Add(new PptxEditPlanProof(operation, elementHash, chart.Binding.PartPath));
                    continue;
                }
                if (!PptxNativeChartLeafCodec.SameTitleBinding(projectedElement.Opaque.NativeChart, chart.Binding))
                    throw new CodecException("presentation_chart_binding_mismatch", $"PPTX edit operation {operation.OperationId} chart-title binding changed after planning.", operation.TargetPartPath);
                if (operation.TextLeafIndex >= (uint)chart.TitleLeaves.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} chart-title leaf index is out of range.", chart.Binding.PartPath);
                if (chart.TitleLeaves[(int)operation.TextLeafIndex].Text != operation.ExpectedValue)
                    throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} chart-title old value does not match the source leaf.", chart.Binding.PartPath);
                proofs.Add(new PptxEditPlanProof(operation, elementHash, chart.Binding.PartPath));
                continue;
            }
            if (LeafKind(operation) is "imageAsset" or "imageSvgAsset")
            {
                var imageProof = LeafKind(operation) == "imageSvgAsset"
                    ? ProveImageSvgReplacement(sourceBytes, slidePart, element, projectedElement, operation, requestedAssets)
                    : ProveImageReplacement(sourceBytes, slidePart, element, projectedElement, operation, requestedAssets);
                proofs.Add(new PptxEditPlanProof(operation, elementHash, operation.SlidePartPath, Image: imageProof));
                continue;
            }
            if (LeafKind(operation) == "deleteElement")
            {
                var deletionProof = ProveElementDeletion(slidePart, tree, element, projectedElement, operation);
                proofs.Add(new PptxEditPlanProof(operation, elementHash, operation.SlidePartPath, Deletion: deletionProof));
                continue;
            }
            if (element is P.Shape shape &&
                projectedElement.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                ((projectedElement.Source.Editable && (LeafKind(operation) is "fillRgb" or "lineRgb" or "leftEmu" or "topEmu" or "widthEmu" or "heightEmu")) ||
                 (projectedElement.Source.TextEditable && LeafKind(operation) == "text" && PptxCodec.SupportsBoundTextLeaf(shape))))
            {
                ProveLeafValue(shape, operation);
            }
            else if (element is P.GraphicFrame table &&
                     projectedElement.ContentCase == PresentationElement.ContentOneofCase.Table &&
                     projectedElement.Source.Editable &&
                     LeafKind(operation) == "tableCellText" &&
                     PptxTableCodec.TryRead(table, out _))
            {
                ProveLeafValue(table, operation);
            }
            else if ((element is P.GraphicFrame || element is P.GroupShape) &&
                     projectedElement.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                     LeafKind(operation) == "nativeText" &&
                     PptxNativeTextLeafCodec.TryResolve(element, operation.TextLeafIndex, out _))
            {
                ProveLeafValue(element, operation);
            }
            else if (element is P.ConnectionShape connector &&
                     projectedElement.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                     (LeafKind(operation) is "lineRgb" or "lineScheme" or "lineWidthEmu") &&
                     PptxNativeObjectCatalog.Classify(connector) == "connector" &&
                     HasSafeNativeConnectorLine(connector, LeafKind(operation)))
            {
                ProveLeafValue(connector, operation);
            }
            else if (element is P.Picture picture &&
                     (projectedElement.ContentCase is PresentationElement.ContentOneofCase.Image or PresentationElement.ContentOneofCase.Opaque) &&
                     projectedElement.Source.Editable &&
                     PptxNativeObjectCatalog.SupportsPlacementEditing(picture) &&
                     IsGeometryLeaf(LeafKind(operation)))
            {
                ProveLeafValue(picture, operation);
            }
            else
            {
                throw new CodecException("unsupported_presentation_edit", $"PPTX edit operation {operation.OperationId} target is not a safely editable shape, table, or picture leaf.", operation.SlidePartPath);
            }
            proofs.Add(new PptxEditPlanProof(operation, elementHash, operation.SlidePartPath));
        }
        return proofs;
    }

    private static PptxXmlPatch[] CompileXmlPatches(byte[] partBytes, IReadOnlyList<PptxEditPlanProof> proofs)
    {
        if (proofs.All(proof => LeafKind(proof.Operation) == "diagramText"))
            return CompileDiagramTextXmlPatches(partBytes, proofs);
        if (proofs.Any(proof => LeafKind(proof.Operation) == "diagramText"))
            throw new CodecException("presentation_edit_plan_scope_violation", "PPTX edit plan mixed SmartArt and non-SmartArt leaves in one mutation part.", proofs[0].MutationPartPath);
        if (proofs.All(proof => LeafKind(proof.Operation) is "chartTitleText" or "chartDataValue"))
        {
            var chartPatches = new List<PptxXmlPatch>();
            var titleProofs = proofs.Where(proof => LeafKind(proof.Operation) == "chartTitleText").ToArray();
            var dataProofs = proofs.Where(proof => LeafKind(proof.Operation) == "chartDataValue").ToArray();
            if (titleProofs.Length > 0) chartPatches.AddRange(CompileChartTitleXmlPatches(partBytes, titleProofs));
            if (dataProofs.Length > 0) chartPatches.AddRange(CompileChartDataXmlPatches(partBytes, dataProofs));
            return OrderedNonOverlapping(chartPatches, proofs[0].MutationPartPath);
        }
        if (proofs.Any(proof => LeafKind(proof.Operation) is "chartTitleText" or "chartDataValue"))
            throw new CodecException("presentation_edit_plan_scope_violation", "PPTX edit plan mixed slide and ChartPart leaves in one mutation part.", proofs[0].MutationPartPath);
        var (xml, bomBytes) = DecodeXml(partBytes);
        var drawingPrefixes = NamespacePattern().Matches(xml)
            .Where(match => match.Groups["uri"].Value == DrawingNamespace)
            .Select(match => match.Groups["prefix"].Value)
            .ToHashSet(StringComparer.Ordinal);
        if (drawingPrefixes.Count == 0)
            throw new CodecException("presentation_edit_target_missing", "PPTX slide XML does not declare the DrawingML namespace.", proofs[0].Operation.SlidePartPath);
        var elements = ShapeElementRanges(xml, "spTree");
        var patches = new List<PptxXmlPatch>();
        foreach (var proof in proofs)
        {
            var operation = proof.Operation;
            var path = ShapeTreePath(operation);
            if (path[0] >= (uint)elements.Count)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw shape-tree index is out of range.", operation.SlidePartPath);
            var range = elements[(int)path[0]];
            for (var depth = 1; depth < path.Count; depth++)
            {
                if (range.LocalName != "grpSp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} shape-tree path crosses a non-group element.", operation.SlidePartPath);
                var groupXml = xml[range.Start..range.End];
                var children = ShapeElementRanges(groupXml, "grpSp");
                if (path[depth] >= (uint)children.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} group-child index is out of range.", operation.SlidePartPath);
                var child = children[(int)path[depth]];
                range = new XmlRange(range.Start + child.Start, range.Start + child.End, child.LocalName);
            }
            var leafKind = LeafKind(operation);
            if (leafKind == "deleteElement")
            {
                if (range.LocalName is not ("sp" or "pic" or "graphicFrame" or "cxnSp" or "grpSp"))
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX element deletion {operation.OperationId} raw target has an unsupported shape-tree type.", operation.SlidePartPath);
                patches.Add(new PptxXmlPatch(operation, range.Start, range.End, string.Empty, proof.SourceElementSha256, proof.MutationPartPath));
                continue;
            }
            if (range.LocalName is not ("sp" or "pic" or "graphicFrame" or "cxnSp" or "grpSp"))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} raw target is not p:sp, p:pic, p:graphicFrame, p:cxnSp, or p:grpSp.", operation.SlidePartPath);
            if (leafKind is not ("text" or "tableCellText" or "nativeText"))
            {
                if (leafKind is "imageAsset" or "imageSvgAsset")
                {
                    patches.Add(CompileImageXmlPatch(xml, range, proof));
                    continue;
                }
                patches.Add(CompileScalarXmlPatch(xml, range, proof));
                continue;
            }
            if ((leafKind == "text" && range.LocalName != "sp") ||
                (leafKind == "tableCellText" && range.LocalName != "graphicFrame") ||
                (leafKind == "nativeText" && range.LocalName is not ("graphicFrame" or "grpSp")))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
            if (leafKind == "tableCellText")
            {
                patches.Add(CompileTableCellTextXmlPatch(xml, range, proof, drawingPrefixes));
                continue;
            }
            if (leafKind == "nativeText")
            {
                patches.Add(CompileNativeTextXmlPatch(xml, range, proof, drawingPrefixes));
                continue;
            }
            var elementXml = xml[range.Start..range.End];
            var leaves = TextLeafPattern().Matches(elementXml)
                .Where(IsNonSelfClosingTextLeaf)
                .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
                .ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
            var leaf = leaves[operation.TextLeafIndex];
            var decoded = DecodeTextLeaf(leaf.Value, leaf.Groups["prefix"].Value);
            if (decoded != operation.ExpectedValue)
                throw new CodecException("presentation_text_precondition_failed", $"PPTX edit operation {operation.OperationId} raw leaf does not match the expected text.", operation.SlidePartPath);
            var open = leaf.Groups["open"].Value;
            if (NeedsPreserve(operation.Value) && !PreserveSpacePattern().IsMatch(open))
                open = open.Insert(open.Length - 1, " xml:space=\"preserve\"");
            var replacement = open + EscapeText(operation.Value) + leaf.Groups["close"].Value;
            patches.Add(new PptxXmlPatch(
                operation,
                range.Start + leaf.Index,
                range.Start + leaf.Index + leaf.Length,
                replacement,
                proof.SourceElementSha256,
                proof.MutationPartPath));
        }
        var ordered = patches.OrderBy(patch => patch.Start).ToArray();
        for (var index = 1; index < ordered.Length; index++)
            if (ordered[index - 1].End > ordered[index].Start)
                throw new CodecException("overlapping_presentation_edit_operations", "PPTX edit plan operations overlap in the source XML.", ordered[index].Operation.SlidePartPath);
        return ordered;
    }

    private static PptxXmlPatch CompileTableCellTextXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var cells = TableCellPattern().Matches(elementXml)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)cells.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw table-cell index is out of range.", operation.SlidePartPath);
        var cell = cells[operation.TextLeafIndex];
        var leaves = TextLeafPattern().Matches(cell.Value)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (leaves.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} table cell does not contain exactly one non-empty text leaf.", operation.SlidePartPath);
        var leaf = leaves[0];
        var prefix = leaf.Groups["prefix"].Value;
        if (DecodeTextLeaf(leaf.Value, prefix) != operation.ExpectedValue)
            throw new CodecException("presentation_text_precondition_failed", $"PPTX edit operation {operation.OperationId} raw table-cell leaf does not match the expected text.", operation.SlidePartPath);
        var open = leaf.Groups["open"].Value;
        if (NeedsPreserve(operation.Value) && !PreserveSpacePattern().IsMatch(open))
            open = open.Insert(open.Length - 1, " xml:space=\"preserve\"");
        var start = elementRange.Start + cell.Index + leaf.Index;
        return new PptxXmlPatch(
            operation,
            start,
            start + leaf.Length,
            open + EscapeText(operation.Value) + leaf.Groups["close"].Value,
            proof.SourceElementSha256,
            proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileNativeTextXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = NativeTextLeafMatches(elementXml, elementRange.LocalName, drawingPrefixes);
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw native text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var prefix = leaf.Groups["prefix"].Value;
        if (DecodeTextLeaf(leaf.Value, prefix) != operation.ExpectedValue)
            throw new CodecException("presentation_text_precondition_failed", $"PPTX edit operation {operation.OperationId} raw native text leaf does not match the expected text.", operation.SlidePartPath);
        var open = leaf.Groups["open"].Value;
        if (NeedsPreserve(operation.Value) && !PreserveSpacePattern().IsMatch(open))
            open = open.Insert(open.Length - 1, " xml:space=\"preserve\"");
        var replacement = ReplaceTextLeaf(leaf, prefix, operation.Value, open);
        return new PptxXmlPatch(
            operation,
            elementRange.Start + leaf.Index,
            elementRange.Start + leaf.Index + leaf.Length,
            replacement,
            proof.SourceElementSha256,
            proof.MutationPartPath);
    }

    private static Match[] NativeTextLeafMatches(
        string elementXml,
        string elementLocalName,
        IReadOnlySet<string> drawingPrefixes)
    {
        var texts = TextLeafPattern().Matches(elementXml)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (elementLocalName != "grpSp") return texts;

        var runs = TextRunPattern().Matches(elementXml).Cast<Match>().ToArray();
        var shapes = ShapeTextPattern().Matches(elementXml).Cast<Match>().ToArray();
        var cells = TableCellPattern().Matches(elementXml).Cast<Match>().ToArray();
        return texts
            .Where(text => runs.Any(run => Contains(run, text)))
            .Where(text => !cells.Any(cell => Contains(cell, text)))
            .Where(text => shapes.Count(shape => Contains(shape, text)) == 1)
            .ToArray();
    }

    private static bool Contains(Match container, Match candidate) =>
        candidate.Index >= container.Index &&
        candidate.Index + candidate.Length <= container.Index + container.Length;

    private static bool IsNonSelfClosingTextLeaf(Match match) =>
        !match.Groups["open"].Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal);

    private static string ReplaceTextLeaf(Match leaf, string prefix, string value, string? openOverride = null)
    {
        var open = openOverride ?? leaf.Groups["open"].Value;
        if (!IsNonSelfClosingTextLeaf(leaf))
        {
            var expanded = open.TrimEnd();
            expanded = expanded[..^2] + ">";
            return expanded + EscapeText(value) + $"</{prefix}:t>";
        }
        return open + EscapeText(value) + leaf.Groups["close"].Value;
    }

    private static PptxXmlPatch[] CompileChartTitleXmlPatches(byte[] partBytes, IReadOnlyList<PptxEditPlanProof> proofs)
    {
        var (xml, _) = DecodeXml(partBytes);
        var drawingPrefixes = NamespacePattern().Matches(xml)
            .Where(match => match.Groups["uri"].Value == DrawingNamespace)
            .Select(match => match.Groups["prefix"].Value)
            .ToHashSet(StringComparer.Ordinal);
        var chartPrefixes = NamespacePattern().Matches(xml)
            .Where(match => match.Groups["uri"].Value == ChartNamespace)
            .Select(match => match.Groups["prefix"].Value)
            .ToHashSet(StringComparer.Ordinal);
        if (drawingPrefixes.Count == 0 || chartPrefixes.Count == 0)
            throw new CodecException("presentation_edit_target_missing", "PPTX ChartPart does not declare the required DrawingML chart namespaces.", proofs[0].MutationPartPath);
        var chartChildren = ShapeElementRanges(xml, "chart");
        var titles = chartChildren.Where(range => range.LocalName == "title").ToArray();
        if (titles.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", "PPTX native chart edit requires one direct chart title.", proofs[0].MutationPartPath);
        var title = titles[0];
        var tx = DirectChildRange(xml, title, "title", "tx", proofs[0].Operation);
        var rich = DirectChildRange(xml, tx, "tx", "rich", proofs[0].Operation);
        var richXml = xml[rich.Start..rich.End];
        var leaves = TextLeafPattern().Matches(richXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        var patches = new List<PptxXmlPatch>();
        foreach (var proof in proofs)
        {
            var operation = proof.Operation;
            if (operation.TextLeafIndex >= (uint)leaves.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw chart-title leaf index is out of range.", proof.MutationPartPath);
            var leaf = leaves[operation.TextLeafIndex];
            var decoded = DecodeTextLeaf(leaf.Value, leaf.Groups["prefix"].Value);
            if (decoded != operation.ExpectedValue)
                throw new CodecException("presentation_text_precondition_failed", $"PPTX edit operation {operation.OperationId} raw chart-title leaf does not match the expected text.", proof.MutationPartPath);
            var open = leaf.Groups["open"].Value;
            if (NeedsPreserve(operation.Value) && !PreserveSpacePattern().IsMatch(open))
                open = open.Insert(open.Length - 1, " xml:space=\"preserve\"");
            patches.Add(new PptxXmlPatch(
                operation,
                rich.Start + leaf.Index,
                rich.Start + leaf.Index + leaf.Length,
                open + EscapeText(operation.Value) + leaf.Groups["close"].Value,
                proof.SourceElementSha256,
                proof.MutationPartPath));
        }
        var ordered = patches.OrderBy(patch => patch.Start).ToArray();
        for (var index = 1; index < ordered.Length; index++)
            if (ordered[index - 1].End > ordered[index].Start)
                throw new CodecException("overlapping_presentation_edit_operations", "PPTX chart-title edit operations overlap in the source XML.", proofs[0].MutationPartPath);
        return ordered;
    }

    private static void ProveLeafValue(OpenXmlElement element, PresentationEditOperation operation)
    {
        var actual = ReadLeafValue(element, operation);
        if (!LeafValuesEqual(actual, operation.ExpectedValue, LeafKind(operation)))
            throw new CodecException(
                "presentation_leaf_precondition_failed",
                $"PPTX edit operation {operation.OperationId} old {LeafKind(operation)} value does not match the source leaf.",
                operation.SlidePartPath);
    }

    private static bool HasSafeNativeConnectorLine(P.ConnectionShape connector, string kind)
    {
        var outline = connector.ShapeProperties?.Elements<A.Outline>().ToArray();
        if (outline is not { Length: 1 }) return false;
        if (kind == "lineWidthEmu")
            return outline[0].Width?.Value is { } width && width is >= 0 and <= 20_116_800;
        var solidFill = outline[0].Elements<A.SolidFill>().ToArray();
        if (solidFill.Length != 1) return false;
        if (solidFill[0].ChildElements.Count != 1) return false;
        if (kind == "lineRgb")
        {
            var colors = solidFill[0].Elements<A.RgbColorModelHex>().ToArray();
            return colors.Length == 1 && colors[0].Val?.Value is { Length: 6 } value && value.All(Uri.IsHexDigit);
        }
        if (kind == "lineScheme")
        {
            var colors = solidFill[0].Elements<A.SchemeColor>().ToArray();
            return colors.Length == 1 && colors[0].Val?.Value is { } value && PptxColor.TrySchemeToken(value, out _);
        }
        return false;
    }

    private static string ReadLeafValue(OpenXmlElement element, PresentationEditOperation operation)
    {
        var kind = LeafKind(operation);
        if (kind == "text")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-leaf index is out of range.", operation.SlidePartPath);
            return leaves[operation.TextLeafIndex].Text;
        }
        if (kind == "tableCellText")
        {
            if (element is not P.GraphicFrame table || !PptxTableCodec.TryRead(table, out var projected))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} tableCellText target is not a supported table.", operation.SlidePartPath);
            var cells = projected.Rows.SelectMany(row => row.Cells).ToArray();
            if (operation.TextLeafIndex >= (uint)cells.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} table-cell index is out of range.", operation.SlidePartPath);
            return cells[operation.TextLeafIndex].Text;
        }
        if (kind == "nativeText")
        {
            if ((element is not P.GraphicFrame && element is not P.GroupShape) ||
                !PptxNativeTextLeafCodec.TryResolve(element, operation.TextLeafIndex, out var leaf))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} nativeText target is not a bounded DrawingML text leaf.", operation.SlidePartPath);
            return leaf.Text;
        }
        var properties = element switch
        {
            P.Shape shape => shape.ShapeProperties,
            P.ConnectionShape connector when kind is "lineRgb" or "lineScheme" or "lineWidthEmu" => connector.ShapeProperties,
            P.Picture picture when IsGeometryLeaf(kind) => picture.ShapeProperties,
            _ => null,
        } ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no p:spPr.", operation.SlidePartPath);
        var transform = properties.Transform2D;
        return kind switch
        {
            "fillRgb" => RequiredLeafValue(PptxColor.SolidRgb(properties.GetFirstChild<A.SolidFill>()), operation),
            "lineRgb" => RequiredLeafValue(PptxColor.SolidRgb(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()), operation),
            "lineScheme" => RequiredLeafValue(NativeSchemeToken(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(), operation), operation),
            "lineWidthEmu" => RequiredLeafValue(properties.GetFirstChild<A.Outline>()?.Width is { } width
                ? width.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : string.Empty, operation),
            "leftEmu" => transform?.Offset?.X?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "topEmu" => transform?.Offset?.Y?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "widthEmu" => transform?.Extents?.Cx?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "heightEmu" => transform?.Extents?.Cy?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported leaf kind {kind}.", operation.SlidePartPath),
        };
    }

    private static string RequiredLeafValue(string value, PresentationEditOperation operation) =>
        string.IsNullOrEmpty(value) ? MissingLeaf(operation) : value;

    private static string NativeSchemeToken(A.SolidFill? fill, PresentationEditOperation operation)
    {
        var scheme = fill?.GetFirstChild<A.SchemeColor>()?.Val?.Value;
        return scheme is { } value && PptxColor.TrySchemeToken(value, out var token)
            ? token
            : MissingLeaf(operation);
    }

    private static string MissingLeaf(PresentationEditOperation operation) =>
        throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no {LeafKind(operation)} leaf.", operation.SlidePartPath);

    private static PptxXmlPatch CompileScalarXmlPatch(string xml, XmlRange elementRange, PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var owner = elementRange.LocalName;
        var properties = DirectChildRange(xml, elementRange, owner, "spPr", operation);
        XmlRange leaf;
        string attribute;
        switch (LeafKind(operation))
        {
            case "fillRgb":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} picture targets do not expose fillRgb.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, "spPr", "solidFill", operation), "solidFill", "srgbClr", operation);
                attribute = "val";
                break;
            case "lineRgb":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineRgb.", operation.SlidePartPath);
                var outline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                leaf = DirectChildRange(xml, DirectChildRange(xml, outline, "ln", "solidFill", operation), "solidFill", "srgbClr", operation);
                attribute = "val";
                break;
            case "lineScheme":
                if (owner != "cxnSp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineScheme.", operation.SlidePartPath);
                var schemeOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                leaf = DirectChildRange(xml, DirectChildRange(xml, schemeOutline, "ln", "solidFill", operation), "solidFill", "schemeClr", operation);
                attribute = "val";
                break;
            case "lineWidthEmu":
                if (owner != "cxnSp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineWidthEmu.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "ln", operation);
                attribute = "w";
                break;
            case "leftEmu":
            case "topEmu":
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, "spPr", "xfrm", operation), "xfrm", "off", operation);
                attribute = LeafKind(operation) == "leftEmu" ? "x" : "y";
                break;
            case "widthEmu":
            case "heightEmu":
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, "spPr", "xfrm", operation), "xfrm", "ext", operation);
                attribute = LeafKind(operation) == "widthEmu" ? "cx" : "cy";
                break;
            default:
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} is not a scalar leaf.", operation.SlidePartPath);
        }
        var fragment = xml[leaf.Start..leaf.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == leaf.LocalName) ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} scalar leaf tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} scalar leaf attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (valueGroup.Value != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw scalar does not match the expected value.", operation.SlidePartPath);
        var start = leaf.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static XmlRange DirectChildRange(
        string xml,
        XmlRange parent,
        string parentLocalName,
        string childLocalName,
        PresentationEditOperation operation)
    {
        var fragment = xml[parent.Start..parent.End];
        var children = ShapeElementRanges(fragment, parentLocalName).Where(child => child.LocalName == childLocalName).ToArray();
        if (children.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct {childLocalName} child under {parentLocalName}.", operation.SlidePartPath);
        var child = children[0];
        return new XmlRange(parent.Start + child.Start, parent.Start + child.End, child.LocalName);
    }

    private static byte[] ApplyPatches(
        byte[] sourcePart,
        IReadOnlyList<PptxXmlPatch> patches,
        ICollection<PresentationEditOperationResult> results)
    {
        var (xml, bomBytes) = DecodeXml(sourcePart);
        var output = xml;
        foreach (var patch in patches.OrderByDescending(patch => patch.Start))
            output = output[..patch.Start] + patch.Replacement + output[patch.End..];
        foreach (var patch in patches)
        {
            var sourceStart = checked((ulong)(bomBytes + StrictUtf8.GetByteCount(xml[..patch.Start])));
            var sourceEnd = checked((ulong)(bomBytes + StrictUtf8.GetByteCount(xml[..patch.End])));
            var priorDelta = patches
                .Where(candidate => candidate.Start < patch.Start)
                .Sum(candidate => candidate.Replacement.Length - (candidate.End - candidate.Start));
            var outputStartCharacter = checked(patch.Start + priorDelta);
            var outputEndCharacter = checked(outputStartCharacter + patch.Replacement.Length);
            var outputEnd = checked((ulong)(bomBytes + StrictUtf8.GetByteCount(output[..outputEndCharacter])));
            var result = new PresentationEditOperationResult
            {
                OperationId = patch.Operation.OperationId,
                SlideId = patch.Operation.SlideId,
                SlidePartPath = patch.Operation.SlidePartPath,
                TargetId = patch.Operation.TargetId,
                ShapeTreeIndex = patch.Operation.ShapeTreeIndex,
                TextLeafIndex = patch.Operation.TextLeafIndex,
                LeafKind = LeafKind(patch.Operation),
                SourceElementSha256 = patch.SourceElementSha256,
                OldValueSha256 = Hash(Encoding.UTF8.GetBytes(patch.Operation.ExpectedValue)),
                NewValueSha256 = Hash(Encoding.UTF8.GetBytes(patch.Operation.Value)),
                SourceStartOffset = sourceStart,
                SourceEndOffset = sourceEnd,
                OutputEndOffset = outputEnd,
                MutationPartPath = patch.MutationPartPath,
            };
            result.ShapeTreePath.Add(ShapeTreePath(patch.Operation));
            results.Add(result);
        }
        var encoded = StrictUtf8.GetBytes(output);
        if (bomBytes == 0) return encoded;
        return StrictUtf8.GetPreamble().Concat(encoded).ToArray();
    }

    private static void VerifyOutput(
        byte[] outputBytes,
        PresentationEditPlanRequest request,
        IReadOnlyList<PresentationEditOperationResult> results,
        EffectiveCodecLimits limits)
    {
        using var stream = new MemoryStream(outputBytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false, new OpenSettings { AutoSave = false });
        var slideByPath = package.PresentationPart!.SlideParts.ToDictionary(PartPath, StringComparer.OrdinalIgnoreCase);
        var resultById = results.ToDictionary(result => result.OperationId, StringComparer.Ordinal);
        foreach (var operation in request.Operations)
        {
            var slidePart = slideByPath[operation.SlidePartPath];
            var tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            if (LeafKind(operation) == "deleteElement")
            {
                var nativeId = operation.ElementDeletion?.ExpectedNativeId ?? 0;
                if (nativeId == 0 || tree.ChildElements.Any(element => PptxElementDeletionCodec.NativeId(element) == nativeId))
                    throw new CodecException("presentation_edit_verification_failed", $"PPTX element deletion {operation.OperationId} did not remove its native target.", operation.SlidePartPath);
                continue;
            }
            var element = ResolveShapeTreeElement(tree, OutputShapeTreePath(operation, request.Operations), operation);
            if (LeafKind(operation) == "diagramText")
            {
                if (!PptxDiagramTextCodec.TryResolveForEditPlan(element, slidePart, out var diagram) ||
                    !PptxNativeObjectCatalog.HasUniqueInboundRelationship(package.PresentationPart!, diagram.Part) ||
                    !diagram.Binding.PartPath.Equals(operation.TargetPartPath, StringComparison.OrdinalIgnoreCase) ||
                    diagram.Binding.RelationshipId != operation.RelationshipId ||
                    operation.TextLeafIndex >= (uint)diagram.Leaves.Count)
                    throw new CodecException("presentation_edit_verification_failed", $"PPTX SmartArt operation {operation.OperationId} did not survive package reopen.", operation.TargetPartPath);
                var leaf = diagram.Leaves[(int)operation.TextLeafIndex];
                if (leaf.ModelId != operation.DiagramModelId || leaf.RunIndex != operation.DiagramRunIndex || leaf.Text != operation.Value)
                    throw new CodecException("presentation_edit_verification_failed", $"PPTX SmartArt text operation {operation.OperationId} did not retain its node/run value.", operation.TargetPartPath);
                resultById[operation.OperationId].OutputElementSha256 = HashElement(element);
                continue;
            }
            if (LeafKind(operation) is "chartTitleText" or "chartDataValue")
            {
                if (!PptxNativeChartLeafCodec.TryResolve(element, slidePart, limits, out var chart) ||
                    !chart.Binding.PartPath.Equals(operation.TargetPartPath, StringComparison.OrdinalIgnoreCase) ||
                    chart.Binding.RelationshipId != operation.RelationshipId)
                    throw new CodecException("presentation_edit_verification_failed", $"PPTX native chart operation {operation.OperationId} did not survive package reopen.", operation.TargetPartPath);
                if (LeafKind(operation) == "chartTitleText")
                {
                    if (operation.TextLeafIndex >= (uint)chart.TitleLeaves.Count || chart.TitleLeaves[(int)operation.TextLeafIndex].Text != operation.Value)
                        throw new CodecException("presentation_edit_verification_failed", $"PPTX chart-title edit operation {operation.OperationId} did not survive package reopen.", operation.TargetPartPath);
                }
                else
                {
                    var point = chart.Data?.Points.SingleOrDefault(candidate =>
                        candidate.Binding.SeriesIndex == operation.ChartSeriesIndex && candidate.Binding.PointIndex == operation.ChartPointIndex);
                    if (point is null || point.Binding.Value != operation.Value || point.Binding.Formula != operation.ChartFormula ||
                        !chart.Binding.EmbeddedPackagePartPath.Equals(operation.EmbeddedPackagePartPath, StringComparison.OrdinalIgnoreCase) ||
                        chart.Binding.EmbeddedPackageRelationshipId != operation.EmbeddedPackageRelationshipId ||
                        !point.Binding.WorksheetPartPath.Equals(operation.EmbeddedWorksheetPartPath, StringComparison.OrdinalIgnoreCase) ||
                        !point.Binding.CellReference.Equals(operation.EmbeddedCellReference, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_edit_verification_failed", $"PPTX chart data operation {operation.OperationId} did not survive package reopen in both cache and workbook.", operation.EmbeddedPackagePartPath);
                }
                resultById[operation.OperationId].OutputElementSha256 = HashElement(element);
                continue;
            }
            if (LeafKind(operation) is "imageAsset" or "imageSvgAsset")
            {
                VerifyImageReplacement(slidePart, element, operation, resultById[operation.OperationId]);
                continue;
            }
            if (element is not P.Shape && element is not P.Picture && element is not P.GraphicFrame && element is not P.GroupShape && element is not P.ConnectionShape)
                throw new CodecException("presentation_edit_verification_failed", "PPTX edited target is no longer a shape, table, group, or picture.", operation.SlidePartPath);
            if (!LeafValuesEqual(ReadLeafValue(element, operation), operation.Value, LeafKind(operation)))
                throw new CodecException("presentation_edit_verification_failed", $"PPTX edit operation {operation.OperationId} did not survive package reopen.", operation.SlidePartPath);
            resultById[operation.OperationId].OutputElementSha256 = HashElement(element);
        }
    }

    internal static byte[] AppendShapeTreeChildren(
        byte[] sourcePart,
        IReadOnlyList<string> childXml,
        string partPath)
    {
        if (childXml.Count == 0) return sourcePart;
        var (xml, bomBytes) = DecodeXml(sourcePart);
        var tokens = XmlTokenPattern().Matches(xml).Cast<Match>().ToArray();
        var treeTokenIndex = Array.FindIndex(tokens, token =>
            !token.Value.StartsWith("</", StringComparison.Ordinal) &&
            !token.Value.EndsWith("/>", StringComparison.Ordinal) &&
            LocalName(token.Value) == "spTree");
        if (treeTokenIndex < 0)
            throw new CodecException("missing_presentation_shape_tree", "PPTX SlidePart XML has no non-empty spTree element.", partPath);

        var depth = 1;
        var insertionIndex = -1;
        for (var index = treeTokenIndex + 1; index < tokens.Length; index++)
        {
            var token = tokens[index];
            var value = token.Value;
            if (value.StartsWith("<!--", StringComparison.Ordinal) ||
                value.StartsWith("<![CDATA[", StringComparison.Ordinal) ||
                value.StartsWith("<?", StringComparison.Ordinal))
                continue;
            var closing = value.StartsWith("</", StringComparison.Ordinal);
            var selfClosing = value.EndsWith("/>", StringComparison.Ordinal);
            if (closing) depth--;
            else if (!selfClosing) depth++;
            if (depth == 0)
            {
                if (!closing || LocalName(value) != "spTree")
                    throw new CodecException("invalid_presentation_shape_tree", "PPTX SlidePart spTree token topology is not safely appendable.", partPath);
                insertionIndex = token.Index;
                break;
            }
        }
        if (insertionIndex < 0)
            throw new CodecException("invalid_presentation_shape_tree", "PPTX SlidePart spTree has no matching closing token.", partPath);

        var output = xml.Insert(insertionIndex, string.Concat(childXml));
        var encoded = StrictUtf8.GetBytes(output);
        if (bomBytes == 0) return encoded;
        return StrictUtf8.GetPreamble().Concat(encoded).ToArray();
    }

    internal static byte[] ReplaceParts(byte[] sourceBytes, IReadOnlyDictionary<string, byte[]> replacements)
        => RewriteParts(
            sourceBytes,
            replacements,
            new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase),
            new HashSet<string>(StringComparer.OrdinalIgnoreCase));

    private static byte[] RewriteParts(
        byte[] sourceBytes,
        IReadOnlyDictionary<string, byte[]> replacements,
        IReadOnlyDictionary<string, byte[]> additions,
        IReadOnlySet<string> removals)
    {
        if (replacements.Keys.Any(removals.Contains) || additions.Keys.Any(removals.Contains) || replacements.Keys.Any(additions.ContainsKey))
            throw new CodecException("presentation_edit_plan_scope_violation", "PPTX edit plan cannot add, replace, and remove the same OPC part.");
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var path in removals.OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                var removed = archive.GetEntry(path) ?? throw new CodecException("presentation_edit_target_missing", $"PPTX part {path} selected for deletion is missing.", path);
                removed.Delete();
            }
            foreach (var (path, data) in replacements)
            {
                var source = archive.GetEntry(path) ?? throw new CodecException("presentation_edit_target_missing", $"PPTX part {path} is missing.", path);
                var timestamp = source.LastWriteTime;
                var attributes = source.ExternalAttributes;
                source.Delete();
                var replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
                replacement.LastWriteTime = timestamp;
                replacement.ExternalAttributes = attributes;
                using var target = replacement.Open();
                target.Write(data);
            }
            var addedPartTimestamp = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
            foreach (var (path, data) in additions.OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase))
            {
                if (archive.GetEntry(path) is not null)
                    throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX edit plan cannot add existing part {path}.", path);
                var addition = archive.CreateEntry(path, CompressionLevel.Optimal);
                addition.LastWriteTime = addedPartTimestamp;
                using var target = addition.Open();
                target.Write(data);
            }
        }
        return stream.ToArray();
    }

    private static IReadOnlyDictionary<string, byte[]> PackageParts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.Where(entry => !entry.FullName.EndsWith('/')).ToDictionary(
            entry => entry.FullName,
            entry => ReadEntry(entry),
            StringComparer.OrdinalIgnoreCase);
    }

    internal static byte[] ReadPart(byte[] bytes, string path)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry(path) ?? throw new CodecException("presentation_edit_target_missing", $"PPTX part {path} is missing.", path);
        return ReadEntry(entry);
    }

    private static byte[] RequiredPart(IReadOnlyDictionary<string, byte[]> parts, string path) =>
        parts.TryGetValue(path, out var bytes)
            ? bytes
            : throw new CodecException("presentation_edit_target_missing", $"PPTX part {path} is missing.", path);

    private static byte[] ReadEntry(ZipArchiveEntry entry)
    {
        using var input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static string[] ChangedParts(
        IReadOnlyDictionary<string, byte[]> sourceParts,
        IReadOnlyDictionary<string, byte[]> outputParts)
    {
        return sourceParts.Keys.Concat(outputParts.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path =>
                !sourceParts.TryGetValue(path, out var left) ||
                !outputParts.TryGetValue(path, out var right) ||
                !left.AsSpan().SequenceEqual(right))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string[] ChangedParts(byte[] sourceBytes, byte[] outputBytes) =>
        ChangedParts(PackageParts(sourceBytes), PackageParts(outputBytes));

    private static (string Xml, int BomBytes) DecodeXml(byte[] bytes)
    {
        var bom = bytes.AsSpan().StartsWith(StrictUtf8.GetPreamble()) ? StrictUtf8.GetPreamble().Length : 0;
        var xml = StrictUtf8.GetString(bytes.AsSpan(bom));
        var declaration = Regex.Match(xml, "^<\\?xml\\s+[^>]*encoding\\s*=\\s*(['\"])(?<encoding>.*?)\\1", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        if (declaration.Success && !declaration.Groups["encoding"].Value.Equals("utf-8", StringComparison.OrdinalIgnoreCase))
            throw new CodecException("unsupported_presentation_xml_encoding", "PPTX edit plan supports UTF-8 SlidePart XML only.");
        return (xml, bom);
    }

    private sealed record XmlRange(int Start, int End, string LocalName);

    private static IReadOnlyList<XmlRange> ShapeElementRanges(string xml, string parentLocalName)
    {
        var tokens = XmlTokenPattern().Matches(xml).Cast<Match>().ToArray();
        var treeTokenIndex = Array.FindIndex(tokens, token => !token.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(token.Value) == parentLocalName);
        if (treeTokenIndex < 0) throw new CodecException("missing_presentation_shape_tree", $"PPTX XML has no {parentLocalName} element.");
        var children = new List<XmlRange>();
        var depth = 0;
        var childStart = -1;
        var childName = string.Empty;
        for (var index = treeTokenIndex + 1; index < tokens.Length; index++)
        {
            var token = tokens[index];
            var value = token.Value;
            if (value.StartsWith("<!--", StringComparison.Ordinal) || value.StartsWith("<![CDATA[", StringComparison.Ordinal) || value.StartsWith("<?", StringComparison.Ordinal)) continue;
            var closing = value.StartsWith("</", StringComparison.Ordinal);
            var selfClosing = value.EndsWith("/>", StringComparison.Ordinal);
            var name = LocalName(value);
            if (closing && depth == 0 && name == parentLocalName) break;
            if (!closing && depth == 0)
            {
                childStart = token.Index;
                childName = name;
                if (selfClosing)
                {
                    children.Add(new XmlRange(childStart, token.Index + token.Length, childName));
                    childStart = -1;
                }
                else depth = 1;
                continue;
            }
            if (!closing && !selfClosing) depth++;
            else if (closing) depth--;
            if (depth == 0 && childStart >= 0)
            {
                children.Add(new XmlRange(childStart, token.Index + token.Length, childName));
                childStart = -1;
            }
        }
        return children.Where(child => child.LocalName is not "nvGrpSpPr" and not "grpSpPr").ToArray();
    }

    private static string LocalName(string tag)
    {
        var match = XmlLocalNamePattern().Match(tag);
        return match.Success ? match.Groups["name"].Value : string.Empty;
    }

    private static string LocalAttributeName(string name)
    {
        var separator = name.IndexOf(':');
        return separator >= 0 ? name[(separator + 1)..] : name;
    }

    private static string DecodeTextLeaf(string xml, string prefix)
    {
        try
        {
            var root = XElement.Parse($"<root xmlns:{prefix}=\"{DrawingNamespace}\">{xml}</root>", LoadOptions.PreserveWhitespace);
            return root.Elements().Single().Value;
        }
        catch (Exception exception) when (exception is InvalidOperationException or System.Xml.XmlException)
        {
            throw new CodecException("invalid_presentation_text_leaf", "PPTX text leaf is not safe XML.", innerException: exception);
        }
    }

    private static string EscapeText(string value) => new XText(value).ToString(SaveOptions.DisableFormatting);
    private static bool NeedsPreserve(string value) => value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
    private static IReadOnlyList<uint> ShapeTreePath(PresentationEditOperation operation) =>
        operation.ShapeTreePath.Count > 0 ? operation.ShapeTreePath : [operation.ShapeTreeIndex];
    private static IReadOnlyList<uint> OutputShapeTreePath(
        PresentationEditOperation operation,
        IEnumerable<PresentationEditOperation> operations)
    {
        var path = ShapeTreePath(operation).ToArray();
        var removedBefore = operations.Count(candidate =>
            LeafKind(candidate) == "deleteElement" &&
            candidate.SlidePartPath.Equals(operation.SlidePartPath, StringComparison.OrdinalIgnoreCase) &&
            ShapeTreePath(candidate)[0] < path[0]);
        path[0] = checked(path[0] - (uint)removedBefore);
        return path;
    }
    private static string LeafKind(PresentationEditOperation operation) =>
        string.IsNullOrEmpty(operation.LeafKind) ? "text" : operation.LeafKind;
    private static string MutationPartPath(PresentationEditOperation operation) =>
        LeafKind(operation) is "chartTitleText" or "chartDataValue" or "diagramText" ? operation.TargetPartPath : operation.SlidePartPath;
    private static bool IsGeometryLeaf(string leafKind) =>
        leafKind is "leftEmu" or "topEmu" or "widthEmu" or "heightEmu";
    private static bool LeafValuesEqual(string left, string right, string leafKind) =>
        leafKind is "fillRgb" or "lineRgb"
            ? PptxColor.Normalize(left) == PptxColor.Normalize(right)
            : leafKind == "lineScheme"
                ? PptxColor.NormalizeScheme(left) == PptxColor.NormalizeScheme(right)
            : left == right;
    private static string ShapeTreePathKey(PresentationEditOperation operation) => string.Join("/", ShapeTreePath(operation));
    private static PresentationElement ResolveProjectedElement(
        IList<PresentationElement> elements,
        IReadOnlyList<uint> path,
        PresentationEditOperation operation)
    {
        PresentationElement? element = null;
        var current = elements;
        for (var depth = 0; depth < path.Count; depth++)
        {
            if (path[depth] >= (uint)current.Count)
                throw new CodecException(
                    "presentation_edit_target_missing",
                    $"PPTX edit operation {operation.OperationId} projected shape-tree path is out of range at depth {depth}.",
                    operation.SlidePartPath);
            element = current[(int)path[depth]];
            if (depth + 1 < path.Count)
            {
                if (element.ContentCase != PresentationElement.ContentOneofCase.Group)
                    throw new CodecException(
                        "presentation_edit_target_mismatch",
                        $"PPTX edit operation {operation.OperationId} projected shape-tree path crosses a non-group element.",
                        operation.SlidePartPath);
                current = element.Group.Children;
            }
        }
        return element ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} has an empty projected shape-tree path.", operation.SlidePartPath);
    }
    private static OpenXmlElement ResolveShapeTreeElement(
        P.ShapeTree shapeTree,
        IReadOnlyList<uint> path,
        PresentationEditOperation operation)
    {
        OpenXmlCompositeElement parent = shapeTree;
        OpenXmlElement? element = null;
        for (var depth = 0; depth < path.Count; depth++)
        {
            var children = ShapeElements(parent);
            if (path[depth] >= (uint)children.Length)
                throw new CodecException(
                    "presentation_edit_target_missing",
                    $"PPTX edit operation {operation.OperationId} shape-tree path is out of range at depth {depth}.",
                    operation.SlidePartPath);
            element = children[path[depth]];
            if (depth + 1 < path.Count)
            {
                if (element is not P.GroupShape group)
                    throw new CodecException(
                        "presentation_edit_target_mismatch",
                        $"PPTX edit operation {operation.OperationId} shape-tree path crosses a non-group element.",
                        operation.SlidePartPath);
                parent = group;
            }
        }
        return element ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} has an empty shape-tree path.", operation.SlidePartPath);
    }
    private static OpenXmlElement[] ShapeElements(OpenXmlCompositeElement owner) =>
        owner.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();
    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static bool IsSha256(string value) => value.Length == 64 && value.All(character => char.IsAsciiHexDigit(character));
}
