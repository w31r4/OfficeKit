using System.IO.Compression;
using System.Globalization;
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
        EffectiveCodecLimits limits,
        PresentationArtifact? validatedSourceProjection = null)
    {
        ValidateRequest(sourceBytes, request, limits);
        var sourceHash = Hash(sourceBytes);
        if (validatedSourceProjection is null)
            _ = PackageGuards.ValidateAndCollectOpaque(sourceBytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        var sourceProjection = validatedSourceProjection ?? PptxCodec.Import(sourceBytes, limits).Artifact.Presentation;
        var proofs = ProveOperations(sourceBytes, request, sourceProjection, limits);
        // Keep the source package indexed by path, but materialize only the
        // parts touched by this edit plan. The previous eager dictionary held
        // every OPC payload at once and then built a second copy for the
        // rewritten package; that made a one-slide edit pay for all media.
        var sourceParts = new LazyPackageParts(sourceBytes);
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
        var changedParts = ChangedPartsStreaming(sourceBytes, outputBytes);
        var expectedParts = patchedParts.Keys.Concat(addedParts.Keys).Concat(removedParts).OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
        if (!changedParts.SequenceEqual(expectedParts, StringComparer.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_edit_plan_scope_violation",
                $"PPTX edit plan changed unexpected OPC parts: {string.Join(", ", changedParts)}.");

        foreach (var (path, expected) in patchedParts)
        {
            var actual = ReadPart(outputBytes, path);
            if (!actual.AsSpan().SequenceEqual(expected))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX edit plan output for {path} differs from the compiled token patch.", path);
        }
        foreach (var (path, expected) in addedParts)
        {
            var actual = ReadPart(outputBytes, path);
            if (!actual.AsSpan().SequenceEqual(expected))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX edit plan added part {path} with unexpected bytes.", path);
        }
        foreach (var path in removedParts)
            if (ContainsPart(outputBytes, path))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX edit plan failed to remove part {path}.", path);

        int sourceValidationWarnings;
        using (PpjBuildProfiler.Measure("post-write.validation"))
        {
            _ = PackageGuards.ValidateAndCollectOpaque(outputBytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
            sourceValidationWarnings = PptxCodec.ValidateEditPlanOutput(sourceBytes, outputBytes, limits);
            VerifyOutput(outputBytes, request, results, limits);
        }

        var result = new PresentationEditPlanResult
        {
            SourceSha256 = sourceHash,
            OutputSha256 = Hash(outputBytes),
        };
        result.ChangedParts.Add(changedParts);
        result.Operations.Add(results.OrderBy(item => item.OperationId, StringComparer.Ordinal));
        var diagnostics = new List<Diagnostic>();
        if (sourceValidationWarnings > 0)
            diagnostics.Add(CodecDiagnostics.Warning(
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
            var customGeometryAdjustmentFormulaLeaf = leafKind == "customGeometryAdjustmentFormula";
            if (customGeometryAdjustmentFormulaLeaf) leafKind = "customGeometryAdjustment";
            var customGeometryGuideFormulaLeaf = leafKind == "customGeometryGuideFormula";
            if (customGeometryGuideFormulaLeaf) leafKind = "customGeometryGuide";
            var flatTextZLeaf = leafKind == "textBodyFlatTextZ";
            if (flatTextZLeaf) leafKind = "textBodyWarpAdjustment";
            if (leafKind is not ("text" or "tableCellText" or "tableHeaderRows" or "tableBandedRows" or "tableBandedColumns" or "tableFirstColumnEmphasis" or "tableLastColumnEmphasis" or "tableLastRow" or "nativeText" or "paragraphAlignment" or "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" or "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" or "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier" or "paragraphMarginLeftEmu" or "paragraphIndentEmu" or "paragraphBulletCharacter" or "paragraphBulletAutoNumberScheme" or "paragraphBulletAutoNumberStartAt" or "paragraphBulletFontFamily" or "paragraphBulletColorRgb" or "paragraphBulletColorScheme" or "paragraphBulletSizePoints" or "paragraphBulletSizePercent" or "paragraphLevel" or "verticalAnchor" or "textBodyAnchorCenter" or "textBodyForceAntiAlias" or "textBodySpaceFirstLastParagraph" or "textBodyCompatibleLineSpacing" or "textBodyFromWordArt" or "textBodyWarpPreset" or "textBodyWarpAdjustment" or "textBodyFlatTextZ" or "textBodyInsetLeftEmu" or "textBodyInsetTopEmu" or "textBodyInsetRightEmu" or "textBodyInsetBottomEmu" or "textBodyWrap" or "textBodyColumnCount" or "textBodyColumnGapEmu" or "textBodyRotationDegrees" or "textBodyVerticalOverflow" or "textBodyHorizontalOverflow" or "textBodyUpright" or "textBodyAutoFit" or "textBodyNormalAutoFitFontScale" or "textBodyNormalAutoFitLineSpacingReduction" or "textBodyColumnDirection" or "textBodyVerticalText" or "fontSizePoints" or "fontFamily" or "fontFamilyEastAsia" or "fontFamilyComplexScript" or "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontColorRgb" or "fontColorScheme" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps" or "fontHighlightRgb" or "fontHighlightScheme" or "textGlowRadiusEmu" or "textDefaultGlowRadiusEmu" or "textDefaultGlowColorRgb" or "textDefaultGlowColorScheme" or "textDefaultGlowOpacityThousandthPercent" or "textDefaultShadowBlurRadiusEmu" or "textDefaultShadowDistanceEmu" or "textDefaultShadowDirectionDegrees" or "textDefaultShadowAlignment" or "textDefaultShadowColorRgb" or "textDefaultShadowColorScheme" or "textDefaultShadowOpacityThousandthPercent" or "textDefaultShadowRotateWithShape" or "textGlowColorRgb" or "textGlowColorScheme" or "textGlowOpacityThousandthPercent" or "textInnerShadowBlurRadiusEmu" or "textDefaultInnerShadowBlurRadiusEmu" or "textDefaultInnerShadowDistanceEmu" or "textDefaultInnerShadowDirectionDegrees" or "textDefaultInnerShadowColorRgb" or "textDefaultInnerShadowColorScheme" or "textDefaultInnerShadowOpacityThousandthPercent" or "textInnerShadowDistanceEmu" or "textInnerShadowDirectionDegrees" or "textInnerShadowColorRgb" or "textInnerShadowColorScheme" or "textInnerShadowOpacityThousandthPercent" or "textReflectionBlurRadiusEmu" or "textDefaultReflectionBlurRadiusEmu" or "textDefaultReflectionDistanceEmu" or "textReflectionStartOpacityThousandthPercent" or "textDefaultReflectionStartOpacityThousandthPercent" or "textReflectionEndOpacityThousandthPercent" or "textDefaultReflectionEndOpacityThousandthPercent" or "textReflectionDistanceEmu" or "textReflectionDirectionDegrees" or "textDefaultReflectionDirectionDegrees" or "textSoftEdgeRadiusEmu" or "textDefaultSoftEdgeRadiusEmu" or "shapeGlowRadiusEmu" or "shapeGlowColorRgb" or "shapeGlowColorScheme" or "shapeGlowOpacityThousandthPercent" or "imageGlowRadiusEmu" or "imageGlowColorRgb" or "imageGlowColorScheme" or "imageGlowOpacityThousandthPercent" or "shapeInnerShadowBlurRadiusEmu" or "shapeInnerShadowDistanceEmu" or "shapeInnerShadowDirectionDegrees" or "shapeInnerShadowColorRgb" or "shapeInnerShadowColorScheme" or "shapeInnerShadowOpacityThousandthPercent" or "imageInnerShadowBlurRadiusEmu" or "imageInnerShadowDistanceEmu" or "imageInnerShadowDirectionDegrees" or "imageInnerShadowColorRgb" or "imageInnerShadowColorScheme" or "imageInnerShadowOpacityThousandthPercent" or "shapeReflectionBlurRadiusEmu" or "shapeReflectionStartOpacityThousandthPercent" or "shapeReflectionEndOpacityThousandthPercent" or "shapeReflectionDistanceEmu" or "shapeReflectionDirectionDegrees" or "imageReflectionBlurRadiusEmu" or "imageReflectionStartOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" or "imageReflectionDirectionDegrees" or "imageReflectionDistanceEmu" or "shapeSoftEdgeRadiusEmu" or "imageSoftEdgeRadiusEmu" or "shape3dExtrusionHeightEmu" or "shape3dDepthEmu" or "shape3dContourWidthEmu" or "shape3dContourRgb" or "shape3dContourColorScheme" or "shape3dExtrusionColorScheme" or "shape3dExtrusionRgb" or "shape3dSceneCameraPreset" or "shape3dSceneCameraZoomThousandthPercent" or "shape3dSceneCameraFov60000" or "shape3dSceneCameraRotationLatitude60000" or "shape3dSceneCameraRotationLongitude60000" or "shape3dSceneCameraRotationRevolution60000" or "shape3dSceneBackdropAnchorXEmu" or "shape3dSceneBackdropAnchorYEmu" or "shape3dSceneBackdropAnchorZEmu" or "shape3dSceneBackdropNormalDxEmu" or "shape3dSceneBackdropNormalDyEmu" or "shape3dSceneBackdropNormalDzEmu" or "shape3dSceneBackdropUpDxEmu" or "shape3dSceneBackdropUpDyEmu" or "shape3dSceneBackdropUpDzEmu" or "shape3dSceneLightRigPreset" or "shape3dSceneLightRigDirection" or "shape3dSceneLightRigRotationLatitude60000" or "shape3dSceneLightRigRotationLongitude60000" or "shape3dSceneLightRigRotationRevolution60000" or "shape3dPresetMaterial" or "shape3dBevelTopWidthEmu" or "shape3dBevelTopHeightEmu" or "shape3dBevelTopPreset" or "shape3dBevelBottomWidthEmu" or "shape3dBevelBottomHeightEmu" or "shape3dBevelBottomPreset" or "fillRgb" or "fillOpacityThousandthPercent" or "shadowOpacityThousandthPercent" or "shadowRotateWithShape" or "imageShadowRotateWithShape" or "imageShadowBlurRadiusEmu" or "imageShadowOpacityThousandthPercent" or "imageShadowDistanceEmu" or "imageShadowDirectionDegrees" or "imageShadowAlignment" or "imageShadowColorRgb" or "imageShadowColorScheme" or "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment" or "shadowColorRgb" or "shadowColorScheme" or "imageOpacityThousandthPercent" or "imageMaskPreset" or "imageMaskAdjustment" or "customGeometryAdjustment" or "customGeometryAdjustmentFormula" or "customGeometryGuide" or "customGeometryGuideFormula" or "customGeometryPathWidth" or "customGeometryPathHeight" or "customGeometryPathArcWidthRadius" or "customGeometryPathArcHeightRadius" or "customGeometryPathArcStartAngle60000" or "customGeometryPathArcSweepAngle60000" or "customGeometryPathLineToX" or "customGeometryPathLineToY" or "customGeometryPathMoveToX" or "customGeometryPathMoveToY" or "customGeometryPathQuadraticEndX" or "customGeometryPathQuadraticEndY" or "customGeometryPathQuadraticControlX" or "customGeometryPathQuadraticControlY" or "customGeometryPathCubicEndX" or "customGeometryPathCubicEndY" or "customGeometryPathCubicControl1X" or "customGeometryPathCubicControl1Y" or "customGeometryPathCubicControl2X" or "customGeometryPathCubicControl2Y" or "customGeometryPathFill" or "customGeometryPathStroke" or "customGeometryPathExtrusionAllowed" or "customGeometryConnectionSiteAngle60000" or "customGeometryConnectionSiteXEmu" or "customGeometryConnectionSiteYEmu" or "customGeometryAdjustmentHandleXEmu" or "customGeometryAdjustmentHandleYEmu" or "customGeometryAdjustmentHandleMinXEmu" or "customGeometryAdjustmentHandleMaxXEmu" or "customGeometryAdjustmentHandleMinYEmu" or "customGeometryAdjustmentHandleMaxYEmu" or "customGeometryAdjustmentHandlePolarMinRadiusEmu" or "customGeometryAdjustmentHandlePolarMaxRadiusEmu" or "customGeometryAdjustmentHandlePolarMinAngle60000" or "customGeometryAdjustmentHandlePolarMaxAngle60000" or "customGeometryAdjustmentHandlePolarXEmu" or "customGeometryAdjustmentHandlePolarYEmu" or "customGeometryTextRectangleLeftEmu" or "customGeometryTextRectangleTopEmu" or "customGeometryTextRectangleRightEmu" or "customGeometryTextRectangleBottomEmu" or "presetGeometryAdjustment" or "fillScheme" or "lineRgb" or "lineScheme" or "lineOpacityThousandthPercent" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength" or "lineWidthEmu" or "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "childLeftEmu" or "childTopEmu" or "childWidthEmu" or "childHeightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical" or "imageAsset" or "imageSvgAsset" or "chartTitleText" or "chartDataCategory" or "chartDataValue" or "chartDataXValue" or "chartDataYValue" or "chartDataBubbleSize" or "diagramText" or "deleteElement"))
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported leaf kind {leafKind}.");
            if (flatTextZLeaf) leafKind = "textBodyFlatTextZ";
            if (customGeometryAdjustmentFormulaLeaf) leafKind = "customGeometryAdjustmentFormula";
            if (customGeometryGuideFormulaLeaf) leafKind = "customGeometryGuideFormula";
            if (!IsSha256(operation.ExpectedSlideSha256) || !IsSha256(operation.ExpectedElementSha256) ||
                !IsSha256(operation.ExpectedSemanticSha256) || !IsSha256(operation.ExpectedTextSha256))
                throw new CodecException("invalid_presentation_edit_precondition", $"PPTX edit operation {operation.OperationId} requires SHA-256 preconditions.");
            if (!Hash(Encoding.UTF8.GetBytes(operation.ExpectedValue)).Equals(operation.ExpectedTextSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_text_hash_mismatch", $"PPTX edit operation {operation.OperationId} expected text does not match expected_text_sha256.");
            if (leafKind is "chartTitleText" or "diagramText" || IsChartDataLeafKind(leafKind))
            {
                if (!DependentXmlPartPathPattern().IsMatch(operation.TargetPartPath) || operation.TargetPartPath.Contains("..", StringComparison.Ordinal) ||
                    !IsSha256(operation.ExpectedTargetPartSha256) || string.IsNullOrWhiteSpace(operation.RelationshipId) || operation.RelationshipId.Length > 255)
                    throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid source-bound dependent-part binding.");
                if (IsChartDataLeafKind(leafKind) && HasEmbeddedWorkbookBinding(operation)) ValidateEmbeddedWorkbookBinding(operation);
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
            if (!IsChartDataLeafKind(leafKind) && (operation.ChartSeriesIndex != 0 || operation.ChartPointIndex != 0))
                throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} cannot attach chart-data indices to {leafKind}.");
            if (IsChartDataLeafKind(leafKind))
            {
                if (leafKind == "chartDataCategory")
                {
                    if (!PpjNativeLeafProjection.ValidTextToken(operation.ExpectedValue) || !PpjNativeLeafProjection.ValidTextToken(operation.Value))
                        throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} chart category must be a bounded text token.");
                }
                else if (!ValidFiniteNumber(operation.ExpectedValue) || !ValidFiniteNumber(operation.Value))
                {
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} chart data value must be a finite numeric token.");
                }
            }
            if (leafKind == "fontSizePoints")
            {
                if (!uint.TryParse(operation.ExpectedValue, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var expectedFontSize) ||
                    !uint.TryParse(operation.Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var requestedFontSize) ||
                    expectedFontSize == 0 || expectedFontSize > 76_800 || requestedFontSize == 0 || requestedFontSize > 76_800)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} font size must be an integer from 1 through 76800 hundredths of a point.");
            }
            if (leafKind is "fontFamily" or "fontFamilyEastAsia" or "fontFamilyComplexScript")
            {
                if (!ValidFontFamilyToken(operation.ExpectedValue) || !ValidFontFamilyToken(operation.Value))
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} font family must be a trimmed literal typeface name of 1 through 255 characters.");
            }
            if (leafKind == "fontLanguage")
            {
                var expected = PptxLanguageTag.Validate(operation.ExpectedValue);
                var requested = PptxLanguageTag.Validate(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text language.");
            }
            if (leafKind is "fontBold" or "fontItalic")
            {
                if (!ValidBooleanToken(operation.ExpectedValue) || !ValidBooleanToken(operation.Value))
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} font style must use canonical boolean tokens 0 or 1.");
            }
            if (leafKind is "tableBandedRows" or "tableBandedColumns" or "tableFirstColumnEmphasis" or "tableLastColumnEmphasis" or "tableLastRow")
            {
                if (!ValidBooleanToken(operation.ExpectedValue) || !ValidBooleanToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical boolean token 0 or 1.");
            }
            if (leafKind == "tableHeaderRows")
            {
                if (!ValidBooleanToken(operation.ExpectedValue) || !ValidBooleanToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} tableHeaderRows must use a changed canonical integer 0 or 1.");
            }
            if (leafKind == "fontUnderline")
            {
                var expected = PptxTextDecoration.NormalizeUnderline(operation.ExpectedValue);
                var requested = PptxTextDecoration.NormalizeUnderline(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its underline.");
            }
            if (leafKind == "fontStrike")
            {
                var expected = PptxTextDecoration.NormalizeStrike(operation.ExpectedValue);
                var requested = PptxTextDecoration.NormalizeStrike(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its strike.");
            }
            if (leafKind is "fontColorRgb" or "fontColorScheme")
            {
                var expected = leafKind == "fontColorRgb" ? PptxColor.Normalize(operation.ExpectedValue) : PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = leafKind == "fontColorRgb" ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its font color.");
            }
            if (leafKind == "fontKerningPoints")
            {
                var expected = PptxTextDecoration.NormalizeKerning(operation.ExpectedValue);
                var requested = PptxTextDecoration.NormalizeKerning(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its font kerning.");
            }
            if (leafKind == "fontBaselinePercent")
            {
                var expected = PptxTextDecoration.NormalizeBaseline(operation.ExpectedValue);
                var requested = PptxTextDecoration.NormalizeBaseline(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its font baseline.");
            }
            if (leafKind == "fontSpacingPoints")
            {
                var expected = PptxTextDecoration.NormalizeSpacing(operation.ExpectedValue);
                var requested = PptxTextDecoration.NormalizeSpacing(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its character spacing.");
            }
            if (leafKind == "fontCaps")
            {
                var expected = PptxTextDecoration.NormalizeCaps(operation.ExpectedValue);
                var requested = PptxTextDecoration.NormalizeCaps(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its capitalization.");
            }
            if (leafKind is "fontHighlightRgb" or "fontHighlightScheme")
            {
                var expected = leafKind == "fontHighlightRgb" ? PptxColor.Normalize(operation.ExpectedValue) : PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = leafKind == "fontHighlightRgb" ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its font highlight.");
            }
            if (leafKind == "textBodyColumnCount")
            {
                var expected = ParseTextBodyColumnCountToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyColumnCountToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body column count.");
            }
            if (leafKind == "textBodyColumnGapEmu")
            {
                var expected = ParseTextBodyColumnGapToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyColumnGapToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body column gap.");
            }
            if (leafKind == "textBodyRotationDegrees")
            {
                var expected = ParseTextBodyRotationToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyRotationToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body rotation.");
            }
            if (leafKind == "textBodyVerticalOverflow")
            {
                var expected = ParseTextBodyVerticalOverflowToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyVerticalOverflowToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body vertical overflow.");
            }
            if (leafKind == "textBodyHorizontalOverflow")
            {
                var expected = ParseTextBodyHorizontalOverflowToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyHorizontalOverflowToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body horizontal overflow.");
            }
            if (leafKind == "textBodyUpright")
            {
                var expected = ParseTextBodyUprightToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyUprightToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body upright flag.");
            }
            if (leafKind == "textBodyAnchorCenter")
            {
                var expected = ParseTextBodyAnchorCenterToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyAnchorCenterToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body anchor-center flag.");
            }
            if (leafKind == "textBodyForceAntiAlias")
            {
                var expected = ParseTextBodyForceAntiAliasToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyForceAntiAliasToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body force anti-alias flag.");
            }
            if (leafKind == "textBodySpaceFirstLastParagraph")
            {
                var expected = ParseTextBodySpaceFirstLastParagraphToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodySpaceFirstLastParagraphToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body first-last paragraph spacing flag.");
            }
            if (leafKind == "textBodyCompatibleLineSpacing")
            {
                var expected = ParseTextBodyCompatibleLineSpacingToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyCompatibleLineSpacingToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body compatible line spacing flag.");
            }
            if (leafKind == "textBodyFromWordArt")
            {
                var expected = ParseTextBodyFromWordArtToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyFromWordArtToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body WordArt marker.");
            }
            if (leafKind == "textBodyWarpPreset")
            {
                var expected = ParseTextBodyWarpPresetToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyWarpPresetToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body warp preset.");
            }
            if (leafKind == "textBodyWarpAdjustment")
            {
                var expected = ParseTextBodyWarpAdjustmentToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyWarpAdjustmentToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body warp adjustment.");
            }
            if (leafKind == "textBodyFlatTextZ")
            {
                var expected = ParseTextBodyFlatTextZToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyFlatTextZToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body flat-text z coordinate.");
            }
            if (leafKind == "textBodyAutoFit")
            {
                var expected = ParseTextBodyAutoFitToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyAutoFitToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body AutoFit mode.");
            }
            if (leafKind is "textBodyNormalAutoFitFontScale" or "textBodyNormalAutoFitLineSpacingReduction")
            {
                var expected = ParseTextBodyNormalAutoFitToken(operation.ExpectedValue, leafKind, operation);
                var requested = ParseTextBodyNormalAutoFitToken(operation.Value, leafKind, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body normal AutoFit percentage.");
            }
            if (leafKind == "textBodyColumnDirection")
            {
                var expected = ParseTextBodyColumnDirectionToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyColumnDirectionToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body column direction.");
            }
            if (leafKind == "textBodyVerticalText")
            {
                var expected = ParseTextBodyVerticalTextToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyVerticalTextToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body vertical text mode.");
            }
            if (leafKind is "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier")
            {
                var expected = ParseParagraphSpacingToken(operation.ExpectedValue, leafKind, operation);
                var requested = ParseParagraphSpacingToken(operation.Value, leafKind, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its paragraph line spacing.");
            }
            if (leafKind is "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" or "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier")
            {
                var expected = ParseParagraphSpacingToken(operation.ExpectedValue, leafKind, operation);
                var requested = ParseParagraphSpacingToken(operation.Value, leafKind, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its paragraph spacing.");
            }
            if (leafKind == "paragraphBulletCharacter")
            {
                if (!TryBulletCharacter(operation.ExpectedValue) || !TryBulletCharacter(operation.Value) || operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} paragraphBulletCharacter must change one Unicode scalar value.");
            }
            if (leafKind == "paragraphBulletAutoNumberScheme")
            {
                if (!PptxBulletCodec.IsAutoNumberScheme(operation.ExpectedValue) ||
                    !PptxBulletCodec.IsAutoNumberScheme(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} paragraphBulletAutoNumberScheme must change a supported auto-number scheme.");
            }
            if (leafKind == "paragraphBulletAutoNumberStartAt")
            {
                var expected = ParseParagraphAutoNumberStartAtToken(operation.ExpectedValue, operation);
                var requested = ParseParagraphAutoNumberStartAtToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its auto-number start.");
            }
            if (leafKind == "paragraphBulletFontFamily")
            {
                if (!ValidFontFamilyToken(operation.ExpectedValue) || !ValidFontFamilyToken(operation.Value) || operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} paragraphBulletFontFamily must change a trimmed literal typeface name.");
            }
            if (leafKind is "paragraphBulletColorRgb" or "paragraphBulletColorScheme")
            {
                var expected = leafKind == "paragraphBulletColorRgb" ? PptxColor.Normalize(operation.ExpectedValue) : PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = leafKind == "paragraphBulletColorRgb" ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its paragraph bullet color.");
            }
            if (leafKind is "paragraphBulletSizePoints" or "paragraphBulletSizePercent")
            {
                var expected = ParseParagraphBulletSizeToken(operation.ExpectedValue, leafKind, operation);
                var requested = ParseParagraphBulletSizeToken(operation.Value, leafKind, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its paragraph bullet size.");
            }
            if (leafKind == "paragraphLevel")
            {
                var expected = ParseParagraphLevelToken(operation.ExpectedValue, operation);
                var requested = ParseParagraphLevelToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its paragraph level.");
            }
            if (leafKind == "rotationDegrees")
            {
                if (!long.TryParse(operation.ExpectedValue, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var expectedRotation) ||
                    !long.TryParse(operation.Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var requestedRotation) ||
                    expectedRotation is < -21_600_000 or > 21_600_000 || requestedRotation is < -21_600_000 or > 21_600_000 ||
                    expectedRotation == requestedRotation)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} rotation must be a changed integer from -21600000 through 21600000 (60000ths of a degree).");
            }
            if (leafKind is "flipHorizontal" or "flipVertical")
            {
                if (!ValidBooleanToken(operation.ExpectedValue) || !ValidBooleanToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical boolean token 0 or 1.");
            }
            if (leafKind == "fillOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} fillOpacityThousandthPercent must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind is "shadowBlurRadiusEmu" or "imageShadowBlurRadiusEmu" or "shadowDistanceEmu" or "imageShadowDistanceEmu" or "shadowDirectionDegrees" or "imageShadowDirectionDegrees" or "shadowAlignment" or "imageShadowAlignment" or "textDefaultShadowBlurRadiusEmu" or "textDefaultShadowDistanceEmu" or "textDefaultShadowDirectionDegrees" or "textDefaultShadowAlignment")
            {
                if (!ValidShadowGeometryToken(leafKind, operation.ExpectedValue) ||
                    !ValidShadowGeometryToken(leafKind, operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native shadow geometry token.");
            }
            if (leafKind is "shadowOpacityThousandthPercent" or "imageShadowOpacityThousandthPercent" or "textDefaultShadowOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shadowOpacityThousandthPercent must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind == "textDefaultShadowRotateWithShape")
            {
                if (!ValidBooleanToken(operation.ExpectedValue) || !ValidBooleanToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} textDefaultShadowRotateWithShape must use a changed canonical boolean token 0 or 1.");
            }
            if (leafKind is "shadowRotateWithShape" or "imageShadowRotateWithShape")
            {
                if (!ValidBooleanToken(operation.ExpectedValue) || !ValidBooleanToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical boolean token 0 or 1.");
            }
            if (leafKind is "shadowColorRgb" or "imageShadowColorRgb" or "textDefaultShadowColorRgb" or "shadowColorScheme" or "imageShadowColorScheme" or "textDefaultShadowColorScheme")
            {
                if (leafKind is "shadowColorRgb" or "imageShadowColorRgb" or "textDefaultShadowColorRgb")
                {
                    var expected = PptxColor.Normalize(operation.ExpectedValue);
                    var requested = PptxColor.Normalize(operation.Value);
                    if (expected == requested)
                        throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its shadow color.");
                }
                else
                {
                    var expected = PptxColor.NormalizeScheme(operation.ExpectedValue);
                    var requested = PptxColor.NormalizeScheme(operation.Value);
                    if (expected == requested)
                        throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its shadow color.");
                }
            }
            if (leafKind == "shape3dContourRgb")
            {
                var expected = PptxColor.Normalize(operation.ExpectedValue);
                var requested = PptxColor.Normalize(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its 3-D contour color.");
            }
            if (leafKind == "shape3dContourColorScheme")
            {
                var expected = PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its 3-D contour theme color.");
            }
            if (leafKind == "shape3dExtrusionRgb")
            {
                var expected = PptxColor.Normalize(operation.ExpectedValue);
                var requested = PptxColor.Normalize(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its 3-D extrusion color.");
            }
            if (leafKind == "shape3dExtrusionColorScheme")
            {
                var expected = PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its 3-D extrusion theme color.");
            }
            if (leafKind == "shape3dSceneCameraPreset")
            {
                if (!PptxShape3DCodec.IsSceneCameraPresetToken(operation.ExpectedValue) ||
                    !PptxShape3DCodec.IsSceneCameraPresetToken(operation.Value) ||
                    string.Equals(operation.ExpectedValue, operation.Value, StringComparison.Ordinal))
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D scene camera preset must use a changed canonical camera-preset token.");
            }
            if (leafKind == "shape3dSceneCameraZoomThousandthPercent")
            {
                if (!PptxShape3DCodec.TryLiteralCoordinate(operation.ExpectedValue, out var expectedCameraZoom) ||
                    !PptxShape3DCodec.TryLiteralCoordinate(operation.Value, out var requestedCameraZoom) ||
                    expectedCameraZoom == requestedCameraZoom)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D scene camera zoom must use a changed canonical non-negative zoom token.");
            }
            if (leafKind == "shape3dSceneCameraFov60000")
            {
                if (!PptxShape3DCodec.TryLiteralCameraFov(operation.ExpectedValue, out var expectedCameraFov) ||
                    !PptxShape3DCodec.TryLiteralCameraFov(operation.Value, out var requestedCameraFov) ||
                    expectedCameraFov == requestedCameraFov)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D scene camera FOV must use a changed canonical positive FOV token below 180 degrees.");
            }
            if (leafKind == "shape3dSceneCameraRotationLatitude60000")
            {
                if (!PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.ExpectedValue, out var expectedCameraLatitude) ||
                    !PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.Value, out var requestedCameraLatitude) ||
                    expectedCameraLatitude == requestedCameraLatitude)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D camera rotation latitude must use a changed canonical non-negative angle token at or below 360 degrees.");
            }
            if (leafKind == "shape3dSceneCameraRotationLongitude60000")
            {
                if (!PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.ExpectedValue, out var expectedCameraLongitude) ||
                    !PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.Value, out var requestedCameraLongitude) ||
                    expectedCameraLongitude == requestedCameraLongitude)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D camera rotation longitude must use a changed canonical non-negative angle token at or below 360 degrees.");
            }
            if (leafKind == "shape3dSceneCameraRotationRevolution60000")
            {
                if (!PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.ExpectedValue, out var expectedCameraRevolution) ||
                    !PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.Value, out var requestedCameraRevolution) ||
                    expectedCameraRevolution == requestedCameraRevolution)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D camera rotation revolution must use a changed canonical non-negative angle token at or below 360 degrees.");
            }
            if (leafKind == "shape3dSceneBackdropAnchorXEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropAnchorX) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropAnchorX) ||
                    expectedBackdropAnchorX == requestedBackdropAnchorX)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop anchor X must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropAnchorYEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropAnchorY) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropAnchorY) ||
                    expectedBackdropAnchorY == requestedBackdropAnchorY)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop anchor Y must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropAnchorZEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropAnchorZ) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropAnchorZ) ||
                    expectedBackdropAnchorZ == requestedBackdropAnchorZ)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop anchor Z must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropNormalDxEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropNormalDx) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropNormalDx) ||
                    expectedBackdropNormalDx == requestedBackdropNormalDx)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop normal X must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropNormalDyEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropNormalDy) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropNormalDy) ||
                    expectedBackdropNormalDy == requestedBackdropNormalDy)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop normal Y must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropNormalDzEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropNormalDz) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropNormalDz) ||
                    expectedBackdropNormalDz == requestedBackdropNormalDz)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop normal Z must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropUpDxEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropUpDx) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropUpDx) ||
                    expectedBackdropUpDx == requestedBackdropUpDx)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop up-vector X must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropUpDyEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropUpDy) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropUpDy) ||
                    expectedBackdropUpDy == requestedBackdropUpDy)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop up-vector Y must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneBackdropUpDzEmu")
            {
                if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.ExpectedValue, out var expectedBackdropUpDz) ||
                    !PptxShape3DCodec.TryLiteralBackdropCoordinate(operation.Value, out var requestedBackdropUpDz) ||
                    expectedBackdropUpDz == requestedBackdropUpDz)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D backdrop up-vector Z must use a changed canonical signed EMU coordinate.");
            }
            if (leafKind == "shape3dSceneLightRigRotationLatitude60000")
            {
                if (!PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.ExpectedValue, out var expectedLatitude) ||
                    !PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.Value, out var requestedLatitude) ||
                    expectedLatitude == requestedLatitude)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D light-rig rotation latitude must use a changed canonical non-negative angle token at or below 360 degrees.");
            }
            if (leafKind == "shape3dSceneLightRigPreset")
            {
                if (!PptxShape3DCodec.IsSceneLightRigPresetToken(operation.ExpectedValue) ||
                    !PptxShape3DCodec.IsSceneLightRigPresetToken(operation.Value) ||
                    string.Equals(operation.ExpectedValue, operation.Value, StringComparison.Ordinal))
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D scene light-rig preset must use a changed canonical light-rig preset token.");
            }
            if (leafKind == "shape3dSceneLightRigDirection")
            {
                if (!PptxShape3DCodec.IsSceneLightRigDirectionToken(operation.ExpectedValue) ||
                    !PptxShape3DCodec.IsSceneLightRigDirectionToken(operation.Value) ||
                    string.Equals(operation.ExpectedValue, operation.Value, StringComparison.Ordinal))
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D scene light-rig direction must use a changed canonical light-rig direction token.");
            }
            if (leafKind == "shape3dSceneLightRigRotationLongitude60000")
            {
                if (!PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.ExpectedValue, out var expectedLongitude) ||
                    !PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.Value, out var requestedLongitude) ||
                    expectedLongitude == requestedLongitude)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D light-rig rotation longitude must use a changed canonical non-negative angle token at or below 360 degrees.");
            }
            if (leafKind == "shape3dSceneLightRigRotationRevolution60000")
            {
                if (!PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.ExpectedValue, out var expectedRevolution) ||
                    !PptxShape3DCodec.TryLiteralSceneRotationAngle(operation.Value, out var requestedRevolution) ||
                    expectedRevolution == requestedRevolution)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shape 3-D light-rig rotation revolution must use a changed canonical non-negative angle token at or below 360 degrees.");
            }
            if (leafKind is "textGlowRadiusEmu" or "textDefaultGlowRadiusEmu")
            {
                if (!ValidGlowRadiusToken(operation.ExpectedValue) || !ValidGlowRadiusToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native glow radius.");
            }
            if (leafKind is "textGlowOpacityThousandthPercent" or "textDefaultGlowOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind is "textGlowColorRgb" or "textDefaultGlowColorRgb" or "textGlowColorScheme" or "textDefaultGlowColorScheme")
            {
                var isRgb = leafKind is "textGlowColorRgb" or "textDefaultGlowColorRgb";
                var expected = isRgb
                    ? PptxColor.Normalize(operation.ExpectedValue)
                    : PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = isRgb
                    ? PptxColor.Normalize(operation.Value)
                    : PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text glow color.");
            }
            if (leafKind is "shapeGlowRadiusEmu" or "imageGlowRadiusEmu")
            {
                if (!ValidGlowRadiusToken(operation.ExpectedValue) || !ValidGlowRadiusToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native glow radius.");
            }
            if (leafKind is "shapeGlowOpacityThousandthPercent" or "imageGlowOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind is "shapeGlowColorRgb" or "shapeGlowColorScheme" or "imageGlowColorRgb" or "imageGlowColorScheme")
            {
                var rgb = leafKind is "shapeGlowColorRgb" or "imageGlowColorRgb";
                var expected = rgb ? PptxColor.Normalize(operation.ExpectedValue) : PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = rgb ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its glow color.");
            }
            if (leafKind is "shapeInnerShadowBlurRadiusEmu" or "shapeInnerShadowDistanceEmu" or "shapeInnerShadowDirectionDegrees" or
                "imageInnerShadowBlurRadiusEmu" or "imageInnerShadowDistanceEmu" or "imageInnerShadowDirectionDegrees")
            {
                if (!ValidInnerShadowGeometryToken(leafKind, operation.ExpectedValue) ||
                    !ValidInnerShadowGeometryToken(leafKind, operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native inner-shadow geometry token.");
            }
            if (leafKind is "shapeInnerShadowOpacityThousandthPercent" or "imageInnerShadowOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind is "shapeInnerShadowColorRgb" or "shapeInnerShadowColorScheme" or "imageInnerShadowColorRgb" or "imageInnerShadowColorScheme")
            {
                var rgb = leafKind is "shapeInnerShadowColorRgb" or "imageInnerShadowColorRgb";
                var expected = rgb ? PptxColor.Normalize(operation.ExpectedValue) : PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = rgb ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its inner-shadow color.");
            }
            if (leafKind is "shapeReflectionBlurRadiusEmu" or "shapeReflectionStartOpacityThousandthPercent" or "shapeReflectionEndOpacityThousandthPercent" or "shapeReflectionDistanceEmu" or "shapeReflectionDirectionDegrees" or
                "imageReflectionBlurRadiusEmu" or "imageReflectionStartOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" or "imageReflectionDistanceEmu" or "imageReflectionDirectionDegrees")
            {
                if (leafKind is "shapeReflectionStartOpacityThousandthPercent" or "shapeReflectionEndOpacityThousandthPercent" or
                    "imageReflectionStartOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent")
                {
                    if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                        operation.ExpectedValue == operation.Value)
                        throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical value from 0 through 100000.");
                }
                else if (!ValidReflectionGeometryToken(leafKind, operation.ExpectedValue) ||
                         !ValidReflectionGeometryToken(leafKind, operation.Value) ||
                         operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native reflection geometry token.");
            }
            if (leafKind is "shapeSoftEdgeRadiusEmu" or "imageSoftEdgeRadiusEmu")
            {
                if (!ValidSoftEdgeRadiusToken(operation.ExpectedValue) ||
                    !ValidSoftEdgeRadiusToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native soft-edge radius.");
            }
            if (leafKind is "textInnerShadowBlurRadiusEmu" or "textDefaultInnerShadowBlurRadiusEmu" or "textInnerShadowDistanceEmu" or "textDefaultInnerShadowDistanceEmu" or "textInnerShadowDirectionDegrees" or "textDefaultInnerShadowDirectionDegrees")
            {
                if (!ValidTextInnerShadowGeometryToken(leafKind, operation.ExpectedValue) ||
                    !ValidTextInnerShadowGeometryToken(leafKind, operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native text inner-shadow geometry token.");
            }
            if (leafKind is "textInnerShadowOpacityThousandthPercent" or "textDefaultInnerShadowOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind is "textInnerShadowColorRgb" or "textDefaultInnerShadowColorRgb" or "textInnerShadowColorScheme" or "textDefaultInnerShadowColorScheme")
            {
                var isRgb = leafKind is "textInnerShadowColorRgb" or "textDefaultInnerShadowColorRgb";
                var expected = isRgb
                    ? PptxColor.Normalize(operation.ExpectedValue)
                    : PptxColor.NormalizeScheme(operation.ExpectedValue);
                var requested = isRgb
                    ? PptxColor.Normalize(operation.Value)
                    : PptxColor.NormalizeScheme(operation.Value);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text inner-shadow color.");
            }
            if (leafKind is "textReflectionBlurRadiusEmu" or "textDefaultReflectionBlurRadiusEmu" or "textDefaultReflectionDistanceEmu" or "textReflectionDistanceEmu" or "textReflectionDirectionDegrees" or "textDefaultReflectionDirectionDegrees")
            {
                if (!ValidTextReflectionGeometryToken(leafKind, operation.ExpectedValue) ||
                    !ValidTextReflectionGeometryToken(leafKind, operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native text reflection geometry token.");
            }
            if (leafKind is "textReflectionStartOpacityThousandthPercent" or "textDefaultReflectionStartOpacityThousandthPercent" or "textReflectionEndOpacityThousandthPercent" or "textDefaultReflectionEndOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind is "textSoftEdgeRadiusEmu" or "textDefaultSoftEdgeRadiusEmu")
            {
                if (!ValidTextSoftEdgeRadiusToken(operation.ExpectedValue) ||
                    !ValidTextSoftEdgeRadiusToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native text soft-edge radius.");
            }
            if (leafKind is "imageOpacityThousandthPercent" or "lineOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind == "imageMaskPreset")
            {
                if (!PptxCustomGeometryCodec.TryPreset(operation.ExpectedValue, out _) ||
                    !PptxCustomGeometryCodec.TryPreset(operation.Value, out _) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} imageMaskPreset must use a changed supported DrawingML preset geometry.");
            }
            if (leafKind == "lineStyle")
            {
                if (!PptxLineStyleCodec.TryPresetDashToken(operation.ExpectedValue, out _) ||
                    !PptxLineStyleCodec.TryPresetDashToken(operation.Value, out _) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} lineStyle must use a changed supported preset style.");
            }
            if (leafKind == "lineCap")
            {
                if (!PptxLineStyleCodec.TryCapToken(operation.ExpectedValue, out _) ||
                    !PptxLineStyleCodec.TryCapToken(operation.Value, out _) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} lineCap must use a changed supported cap.");
            }
            if (leafKind == "lineJoin")
            {
                if (!PptxLineStyleCodec.TryJoinToken(operation.ExpectedValue, out _) ||
                    !PptxLineStyleCodec.TryJoinToken(operation.Value, out _) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} lineJoin must use a changed supported join.");
            }
            if (leafKind is "lineStartArrow" or "lineEndArrow")
            {
                if (!PptxLineStyleCodec.TryArrowTypeToken(operation.ExpectedValue, out _) ||
                    !PptxLineStyleCodec.TryArrowTypeToken(operation.Value, out _) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed supported arrow type.");
            }
            if (leafKind is "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength")
            {
                if (!PptxLineStyleCodec.TryArrowSizeToken(operation.ExpectedValue) ||
                    !PptxLineStyleCodec.TryArrowSizeToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed sm/med/lg arrow size.");
            }
            if (leafKind is not ("chartTitleText" or "diagramText") && !IsChartDataLeafKind(leafKind) &&
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
            if (leafKind == "fillScheme")
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
                    (leafKind is "widthEmu" or "heightEmu" or "childWidthEmu" or "childHeightEmu" && requested <= 0))
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
            if (LeafKind(operation) is "chartTitleText" || IsChartDataLeafKind(LeafKind(operation)))
            {
                if (element is not P.GraphicFrame || projectedElement.ContentCase != PresentationElement.ContentOneofCase.Opaque ||
                    projectedElement.Opaque.NativeChart is null ||
                    !PptxNativeChartLeafCodec.TryResolve(element, slidePart, limits, out var chart) ||
                    !PptxNativeObjectCatalog.HasUniqueInboundRelationship(presentationPart, chart.Part) ||
                    !operation.TargetPartPath.Equals(chart.Binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
                    !operation.ExpectedTargetPartSha256.Equals(chart.Binding.SourceSha256, StringComparison.OrdinalIgnoreCase) ||
                    operation.RelationshipId != chart.Binding.RelationshipId)
                    throw new CodecException("presentation_chart_binding_mismatch", $"PPTX edit operation {operation.OperationId} no longer resolves to its unique source-bound ChartPart.", operation.SlidePartPath);
                if (IsChartDataLeafKind(LeafKind(operation)))
                {
                    if (!PptxNativeChartLeafCodec.SameBinding(projectedElement.Opaque.NativeChart, chart.Binding) || chart.Data is null)
                        throw new CodecException("presentation_chart_data_binding_mismatch", $"PPTX edit operation {operation.OperationId} no longer resolves to its unique source-bound embedded workbook.", operation.TargetPartPath);
                    var point = chart.Data.Points.SingleOrDefault(candidate =>
                        candidate.Binding.SeriesIndex == operation.ChartSeriesIndex &&
                        candidate.Binding.PointIndex == operation.ChartPointIndex &&
                        ChartDataChannel(LeafKind(operation)) == NativeChartDataChannel(candidate.Binding));
                    if (point is null || point.Binding.Value != operation.ExpectedValue || point.Binding.Formula != operation.ChartFormula ||
                        !point.Binding.WorksheetPartPath.Equals(operation.EmbeddedWorksheetPartPath, StringComparison.OrdinalIgnoreCase) ||
                        !point.Binding.WorksheetSourceSha256.Equals(operation.ExpectedEmbeddedWorksheetSha256, StringComparison.OrdinalIgnoreCase) ||
                        !point.Binding.CellReference.Equals(operation.EmbeddedCellReference, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} chart data point no longer matches its cache/workbook binding.", operation.TargetPartPath);
                    if (string.IsNullOrEmpty(operation.EmbeddedPackagePartPath))
                    {
                        if (!string.IsNullOrEmpty(chart.Binding.EmbeddedPackagePartPath) ||
                            !string.IsNullOrEmpty(chart.Binding.EmbeddedPackageSourceSha256) ||
                            !string.IsNullOrEmpty(chart.Binding.EmbeddedPackageRelationshipId) ||
                            !string.IsNullOrEmpty(point.Binding.Formula) ||
                            !string.IsNullOrEmpty(point.Binding.WorksheetPartPath) ||
                            !string.IsNullOrEmpty(point.Binding.WorksheetSourceSha256) ||
                            !string.IsNullOrEmpty(point.Binding.WorksheetName) ||
                            !string.IsNullOrEmpty(point.Binding.CellReference))
                            throw new CodecException("presentation_chart_data_binding_mismatch", $"PPTX edit operation {operation.OperationId} literal chart cache unexpectedly carries an embedded-workbook binding.", operation.TargetPartPath);
                    }
                    else if (chart.Data.Part is null || chart.Data.PackageBytes is null ||
                             !PptxNativeObjectCatalog.HasUniqueInboundRelationship(presentationPart, chart.Data.Part) ||
                             !operation.EmbeddedPackagePartPath.Equals(chart.Binding.EmbeddedPackagePartPath, StringComparison.OrdinalIgnoreCase) ||
                             !operation.ExpectedEmbeddedPackageSha256.Equals(chart.Binding.EmbeddedPackageSourceSha256, StringComparison.OrdinalIgnoreCase) ||
                             operation.EmbeddedPackageRelationshipId != chart.Binding.EmbeddedPackageRelationshipId)
                        throw new CodecException("presentation_chart_data_binding_mismatch", $"PPTX edit operation {operation.OperationId} no longer resolves to its unique source-bound embedded workbook.", operation.TargetPartPath);
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
            if (element is P.Shape defaultTextShape &&
                projectedElement.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                projectedElement.Source.TextEditable &&
                LeafKind(operation) is "textDefaultGlowColorRgb" or "textDefaultGlowColorScheme" or "textDefaultGlowOpacityThousandthPercent" or "textDefaultShadowBlurRadiusEmu" or "textDefaultShadowDistanceEmu" or "textDefaultShadowDirectionDegrees" or "textDefaultShadowAlignment" or "textDefaultShadowColorRgb" or "textDefaultShadowColorScheme" or "textDefaultShadowOpacityThousandthPercent" or "textDefaultShadowRotateWithShape" or "textDefaultInnerShadowBlurRadiusEmu" or "textDefaultInnerShadowDistanceEmu" or "textDefaultInnerShadowDirectionDegrees" or "textDefaultInnerShadowColorRgb" or "textDefaultInnerShadowColorScheme" or "textDefaultInnerShadowOpacityThousandthPercent" or "textDefaultReflectionBlurRadiusEmu" or "textDefaultReflectionDistanceEmu" or "textDefaultReflectionStartOpacityThousandthPercent" or "textDefaultReflectionEndOpacityThousandthPercent" or "textDefaultReflectionDirectionDegrees")
            {
                ProveLeafValue(defaultTextShape, operation);
                proofs.Add(new PptxEditPlanProof(operation, elementHash, operation.SlidePartPath));
                continue;
            }
            if (element is P.Shape shape &&
                projectedElement.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                 ((projectedElement.Source.Editable &&
                   (LeafKind(operation) is "fillRgb" or "fillScheme" or "fillOpacityThousandthPercent" or "shadowOpacityThousandthPercent" or "shadowColorRgb" or "shadowColorScheme" or "lineRgb" or "lineScheme" or "lineWidthEmu" or "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical" ||
                    LeafKind(operation) == "customGeometryAdjustment" && HasSafeNativeCustomGeometryAdjustment(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentFormula" && HasSafeNativeCustomGeometryAdjustmentFormula(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryGuide" && HasSafeNativeCustomGeometryGuide(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryGuideFormula" && HasSafeNativeCustomGeometryGuideFormula(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathWidth" && HasSafeNativeCustomGeometryPathWidth(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathHeight" && HasSafeNativeCustomGeometryPathHeight(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathArcWidthRadius" && HasSafeNativeCustomGeometryPathArcWidthRadius(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathArcHeightRadius" && HasSafeNativeCustomGeometryPathArcHeightRadius(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathArcStartAngle60000" && HasSafeNativeCustomGeometryPathArcStartAngle(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathArcSweepAngle60000" && HasSafeNativeCustomGeometryPathArcSweepAngle(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathLineToX" && HasSafeNativeCustomGeometryPathLineToX(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathLineToY" && HasSafeNativeCustomGeometryPathLineToY(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathMoveToX" && HasSafeNativeCustomGeometryPathMoveToX(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathMoveToY" && HasSafeNativeCustomGeometryPathMoveToY(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathQuadraticEndX" && HasSafeNativeCustomGeometryPathQuadraticEndX(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathQuadraticEndY" && HasSafeNativeCustomGeometryPathQuadraticEndY(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathQuadraticControlX" && HasSafeNativeCustomGeometryPathQuadraticControlX(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathQuadraticControlY" && HasSafeNativeCustomGeometryPathQuadraticControlY(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathCubicEndX" && HasSafeNativeCustomGeometryPathCubicEndX(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathCubicEndY" && HasSafeNativeCustomGeometryPathCubicEndY(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathCubicControl1X" && HasSafeNativeCustomGeometryPathCubicControl1X(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathCubicControl1Y" && HasSafeNativeCustomGeometryPathCubicControl1Y(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathCubicControl2X" && HasSafeNativeCustomGeometryPathCubicControl2X(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathCubicControl2Y" && HasSafeNativeCustomGeometryPathCubicControl2Y(shape, operation.NativeLeafIndex, operation.TextLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathFill" && HasSafeNativeCustomGeometryPathFill(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathStroke" && HasSafeNativeCustomGeometryPathStroke(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryPathExtrusionAllowed" && HasSafeNativeCustomGeometryPathExtrusionAllowed(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryConnectionSiteAngle60000" && HasSafeNativeCustomGeometryConnectionSiteAngle(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryConnectionSiteXEmu" && HasSafeNativeCustomGeometryConnectionSiteX(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryConnectionSiteYEmu" && HasSafeNativeCustomGeometryConnectionSiteY(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandleXEmu" && HasSafeNativeCustomGeometryAdjustmentHandleX(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandleYEmu" && HasSafeNativeCustomGeometryAdjustmentHandleY(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandleMinXEmu" && HasSafeNativeCustomGeometryAdjustmentHandleMinX(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandleMaxXEmu" && HasSafeNativeCustomGeometryAdjustmentHandleMaxX(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandleMinYEmu" && HasSafeNativeCustomGeometryAdjustmentHandleMinY(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandleMaxYEmu" && HasSafeNativeCustomGeometryAdjustmentHandleMaxY(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandlePolarMinRadiusEmu" && HasSafeNativeCustomGeometryAdjustmentHandlePolarMinRadius(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandlePolarMaxRadiusEmu" && HasSafeNativeCustomGeometryAdjustmentHandlePolarMaxRadius(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandlePolarMinAngle60000" && HasSafeNativeCustomGeometryAdjustmentHandlePolarMinAngle(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandlePolarMaxAngle60000" && HasSafeNativeCustomGeometryAdjustmentHandlePolarMaxAngle(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandlePolarXEmu" && HasSafeNativeCustomGeometryAdjustmentHandlePolarX(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryAdjustmentHandlePolarYEmu" && HasSafeNativeCustomGeometryAdjustmentHandlePolarY(shape, operation.NativeLeafIndex) ||
                    LeafKind(operation) == "customGeometryTextRectangleLeftEmu" && HasSafeNativeCustomGeometryTextRectangleLeft(shape) ||
                    LeafKind(operation) == "customGeometryTextRectangleTopEmu" && HasSafeNativeCustomGeometryTextRectangleTop(shape) ||
                    LeafKind(operation) == "customGeometryTextRectangleRightEmu" && HasSafeNativeCustomGeometryTextRectangleRight(shape) ||
                    LeafKind(operation) == "customGeometryTextRectangleBottomEmu" && HasSafeNativeCustomGeometryTextRectangleBottom(shape) ||
                    LeafKind(operation) == "presetGeometryAdjustment" && HasSafeNativePresetGeometryAdjustment(shape, operation.NativeLeafIndex) ||
                    (LeafKind(operation) is "lineStyle" or "lineOpacityThousandthPercent" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength") && HasSafeNativeShapeStyle(shape, LeafKind(operation)) ||
                   (LeafKind(operation) is "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment" or "shadowRotateWithShape") && shape.ShapeProperties is not null && HasSafeNativeShadowGeometry(shape.ShapeProperties, LeafKind(operation))) ||
                   (LeafKind(operation) is "shapeGlowRadiusEmu" or "shapeGlowColorRgb" or "shapeGlowColorScheme" or "shapeGlowOpacityThousandthPercent") && shape.ShapeProperties is not null && HasSafeNativeShapeGlow(shape.ShapeProperties, LeafKind(operation)) ||
                   (LeafKind(operation) is "shapeInnerShadowBlurRadiusEmu" or "shapeInnerShadowDistanceEmu" or "shapeInnerShadowDirectionDegrees" or "shapeInnerShadowColorRgb" or "shapeInnerShadowColorScheme" or "shapeInnerShadowOpacityThousandthPercent") && shape.ShapeProperties is not null && HasSafeNativeShapeInnerShadow(shape.ShapeProperties, LeafKind(operation)) ||
                   (LeafKind(operation) is "shapeReflectionBlurRadiusEmu" or "shapeReflectionStartOpacityThousandthPercent" or "shapeReflectionEndOpacityThousandthPercent" or "shapeReflectionDistanceEmu" or "shapeReflectionDirectionDegrees") && shape.ShapeProperties is not null && HasSafeNativeShapeReflection(shape.ShapeProperties, LeafKind(operation))) ||
                   (LeafKind(operation) == "shapeSoftEdgeRadiusEmu" && shape.ShapeProperties is not null && HasSafeNativeShapeSoftEdge(shape.ShapeProperties)) ||
                   (LeafKind(operation) == "shape3dExtrusionHeightEmu" && HasSafeNativeShape3dExtrusionHeight(shape)) ||
                   (LeafKind(operation) == "shape3dDepthEmu" && HasSafeNativeShape3dDepth(shape)) ||
                   (LeafKind(operation) == "shape3dContourWidthEmu" && HasSafeNativeShape3dContourWidth(shape)) ||
                   (LeafKind(operation) == "shape3dContourRgb" && HasSafeNativeShape3dContourRgb(shape)) ||
                   (LeafKind(operation) == "shape3dContourColorScheme" && HasSafeNativeShape3dContourColorScheme(shape)) ||
                   (LeafKind(operation) == "shape3dExtrusionRgb" && HasSafeNativeShape3dExtrusionRgb(shape)) ||
                   (LeafKind(operation) == "shape3dExtrusionColorScheme" && HasSafeNativeShape3dExtrusionColorScheme(shape)) ||
                   (LeafKind(operation) == "shape3dSceneCameraPreset" && HasSafeNativeShape3dSceneCameraPreset(shape)) ||
                   (LeafKind(operation) == "shape3dSceneCameraZoomThousandthPercent" && HasSafeNativeShape3dSceneCameraZoom(shape)) ||
                   (LeafKind(operation) == "shape3dSceneCameraFov60000" && HasSafeNativeShape3dSceneCameraFov(shape)) ||
                   (LeafKind(operation) == "shape3dSceneCameraRotationLatitude60000" && HasSafeNativeShape3dSceneCameraRotationLatitude(shape)) ||
                   (LeafKind(operation) == "shape3dSceneCameraRotationLongitude60000" && HasSafeNativeShape3dSceneCameraRotationLongitude(shape)) ||
                   (LeafKind(operation) == "shape3dSceneCameraRotationRevolution60000" && HasSafeNativeShape3dSceneCameraRotationRevolution(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropAnchorXEmu" && HasSafeNativeShape3dSceneBackdropAnchorX(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropAnchorYEmu" && HasSafeNativeShape3dSceneBackdropAnchorY(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropAnchorZEmu" && HasSafeNativeShape3dSceneBackdropAnchorZ(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropNormalDxEmu" && HasSafeNativeShape3dSceneBackdropNormalDx(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropNormalDyEmu" && HasSafeNativeShape3dSceneBackdropNormalDy(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropNormalDzEmu" && HasSafeNativeShape3dSceneBackdropNormalDz(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropUpDxEmu" && HasSafeNativeShape3dSceneBackdropUpDx(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropUpDyEmu" && HasSafeNativeShape3dSceneBackdropUpDy(shape)) ||
                   (LeafKind(operation) == "shape3dSceneBackdropUpDzEmu" && HasSafeNativeShape3dSceneBackdropUpDz(shape)) ||
                   (LeafKind(operation) == "shape3dSceneLightRigPreset" && HasSafeNativeShape3dSceneLightRigPreset(shape)) ||
                   (LeafKind(operation) == "shape3dSceneLightRigDirection" && HasSafeNativeShape3dSceneLightRigDirection(shape)) ||
                   (LeafKind(operation) == "shape3dSceneLightRigRotationLatitude60000" && HasSafeNativeShape3dSceneLightRigRotationLatitude(shape)) ||
                   (LeafKind(operation) == "shape3dSceneLightRigRotationLongitude60000" && HasSafeNativeShape3dSceneLightRigRotationLongitude(shape)) ||
                   (LeafKind(operation) == "shape3dSceneLightRigRotationRevolution60000" && HasSafeNativeShape3dSceneLightRigRotationRevolution(shape)) ||
                   (LeafKind(operation) == "shape3dPresetMaterial" && HasSafeNativeShape3dPresetMaterial(shape)) ||
                   (LeafKind(operation) == "shape3dBevelTopWidthEmu" && HasSafeNativeShape3dBevelTopWidth(shape)) ||
                   (LeafKind(operation) == "shape3dBevelTopHeightEmu" && HasSafeNativeShape3dBevelTopHeight(shape)) ||
                   (LeafKind(operation) == "shape3dBevelTopPreset" && HasSafeNativeShape3dBevelTopPreset(shape)) ||
                   (LeafKind(operation) == "shape3dBevelBottomWidthEmu" && HasSafeNativeShape3dBevelBottomWidth(shape)) ||
                   (LeafKind(operation) == "shape3dBevelBottomHeightEmu" && HasSafeNativeShape3dBevelBottomHeight(shape)) ||
                   (LeafKind(operation) == "shape3dBevelBottomPreset" && HasSafeNativeShape3dBevelBottomPreset(shape)) ||
                 (!projectedElement.Source.Editable &&
                  LeafKind(operation) == "presetGeometryAdjustment" &&
                  HasSafeNativePresetGeometryAdjustment(shape, operation.NativeLeafIndex)) ||
                 (!projectedElement.Source.Editable && LeafKind(operation) is ("fillRgb" or "fillOpacityThousandthPercent" or "shadowOpacityThousandthPercent" or "shadowColorRgb" or "shadowColorScheme" or "fillScheme" or "lineRgb" or "lineScheme" or "lineOpacityThousandthPercent" or "lineWidthEmu" or "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment" or "shadowRotateWithShape") && HasSafeNativeShapeStyle(shape, LeafKind(operation))) ||
                   (projectedElement.Source.TextEditable && LeafKind(operation) is ("text" or "paragraphAlignment" or "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" or "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" or "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier" or "paragraphMarginLeftEmu" or "paragraphIndentEmu" or "paragraphBulletCharacter" or "paragraphBulletAutoNumberScheme" or "paragraphBulletAutoNumberStartAt" or "paragraphBulletFontFamily" or "paragraphBulletColorRgb" or "paragraphBulletColorScheme" or "paragraphBulletSizePoints" or "paragraphBulletSizePercent" or "paragraphLevel" or "verticalAnchor" or "textBodyInsetLeftEmu" or "textBodyInsetTopEmu" or "textBodyInsetRightEmu" or "textBodyInsetBottomEmu" or "textBodyWrap" or "textBodyColumnCount" or "textBodyAutoFit" or "textBodyNormalAutoFitFontScale" or "textBodyNormalAutoFitLineSpacingReduction" or "textBodyColumnDirection" or "textBodyVerticalText" or "fontSizePoints" or "fontFamily" or "fontFamilyEastAsia" or "fontFamilyComplexScript" or "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontColorRgb" or "fontColorScheme" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps" or "fontHighlightRgb" or "fontHighlightScheme" or "textGlowRadiusEmu" or "textDefaultGlowRadiusEmu" or "textGlowColorRgb" or "textGlowColorScheme" or "textGlowOpacityThousandthPercent" or "textInnerShadowBlurRadiusEmu" or "textInnerShadowDistanceEmu" or "textInnerShadowDirectionDegrees" or "textInnerShadowColorRgb" or "textInnerShadowColorScheme" or "textInnerShadowOpacityThousandthPercent" or "textReflectionBlurRadiusEmu" or "textDefaultReflectionBlurRadiusEmu" or "textDefaultReflectionDistanceEmu" or "textReflectionStartOpacityThousandthPercent" or "textDefaultReflectionStartOpacityThousandthPercent" or "textReflectionEndOpacityThousandthPercent" or "textDefaultReflectionEndOpacityThousandthPercent" or "textReflectionDistanceEmu" or "textReflectionDirectionDegrees" or "textDefaultReflectionDirectionDegrees" or "textSoftEdgeRadiusEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical") && PptxCodec.SupportsBoundTextLeaf(shape))))
            {
                ProveLeafValue(shape, operation);
            }
            else if (element is P.Shape bodyGapShape &&
                     projectedElement.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                     projectedElement.Source.TextEditable &&
                     (LeafKind(operation) is "textBodyColumnGapEmu" or "textBodyRotationDegrees" or "textBodyVerticalOverflow" or "textBodyHorizontalOverflow" or "textBodyUpright" or "textBodyAnchorCenter" or "textBodyForceAntiAlias" or "textBodySpaceFirstLastParagraph" or "textBodyCompatibleLineSpacing" or "textBodyFromWordArt" or "textBodyWarpPreset" or "textBodyWarpAdjustment" or "textBodyFlatTextZ") &&
                     PptxCodec.SupportsBoundTextLeaf(bodyGapShape))
            {
                ProveLeafValue(bodyGapShape, operation);
            }
            else if (element is P.GroupShape groupGeometry &&
                     projectedElement.ContentCase == PresentationElement.ContentOneofCase.Group &&
                     projectedElement.Source.Editable &&
                     PptxNativeObjectCatalog.SupportsPlacementEditing(groupGeometry) &&
                     (IsGeometryLeaf(LeafKind(operation)) || IsGroupChildGeometryLeaf(LeafKind(operation))))
            {
                ProveLeafValue(groupGeometry, operation);
            }
            else if (element is P.GraphicFrame table &&
                     projectedElement.ContentCase == PresentationElement.ContentOneofCase.Table &&
                     projectedElement.Source.Editable &&
                     LeafKind(operation) is "tableCellText" or "tableHeaderRows" or "tableBandedRows" or "tableBandedColumns" or "tableFirstColumnEmphasis" or "tableLastColumnEmphasis" or "tableLastRow" &&
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
            else if (element is P.GroupShape group &&
                     (projectedElement.ContentCase is PresentationElement.ContentOneofCase.Opaque or PresentationElement.ContentOneofCase.Group) &&
                     (LeafKind(operation) is "fillRgb" or "fillScheme" or "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu") &&
                     PptxNativeStyleLeafCodec.TryResolve(group, operation.NativeLeafIndex, out var styleLeaf) &&
                     styleLeaf.Kind == LeafKind(operation))
            {
                ProveLeafValue(group, operation);
            }
            else if (element is P.ConnectionShape connector &&
                     (projectedElement.ContentCase is PresentationElement.ContentOneofCase.Opaque or PresentationElement.ContentOneofCase.Connector) &&
                     (LeafKind(operation) is "lineRgb" or "lineScheme" or "lineOpacityThousandthPercent" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength" or "lineWidthEmu") &&
                     PptxNativeObjectCatalog.Classify(connector) == "connector" &&
                     HasSafeNativeConnectorLine(connector, LeafKind(operation)))
            {
                ProveLeafValue(connector, operation);
            }
            else if (element is P.Picture picture &&
                     (projectedElement.ContentCase is PresentationElement.ContentOneofCase.Image or PresentationElement.ContentOneofCase.Opaque) &&
                     projectedElement.Source.Editable &&
                     PptxNativeObjectCatalog.SupportsPlacementEditing(picture) &&
                     (IsGeometryLeaf(LeafKind(operation)) ||
                      LeafKind(operation) == "imageOpacityThousandthPercent" && HasSafeNativePictureOpacity(picture) ||
                      LeafKind(operation) == "imageMaskPreset" && HasSafeNativePictureMask(picture) ||
                      LeafKind(operation) == "imageMaskAdjustment" && HasSafeNativePictureMaskAdjustment(picture, operation.NativeLeafIndex) ||
                      LeafKind(operation) is "imageGlowRadiusEmu" or "imageGlowColorRgb" or "imageGlowColorScheme" or "imageGlowOpacityThousandthPercent" && HasSafeNativePictureGlow(picture, LeafKind(operation)) ||
                      LeafKind(operation) == "imageShadowRotateWithShape" && HasSafeNativePictureShadow(picture) ||
                      LeafKind(operation) is "imageShadowBlurRadiusEmu" or "imageShadowDistanceEmu" or "imageShadowDirectionDegrees" or "imageShadowAlignment" && HasSafeNativePictureShadowGeometry(picture, LeafKind(operation)) ||
                      LeafKind(operation) is "imageShadowColorRgb" or "imageShadowColorScheme" && HasSafeNativePictureShadowColor(picture, LeafKind(operation)) ||
                      LeafKind(operation) == "imageShadowOpacityThousandthPercent" && HasSafeNativePictureShadowOpacity(picture) ||
                      LeafKind(operation) == "shape3dExtrusionHeightEmu" && HasSafeNativePictureShape3dExtrusionHeight(picture) ||
                      LeafKind(operation) == "shape3dContourWidthEmu" && HasSafeNativePictureShape3dContourWidth(picture) ||
                      LeafKind(operation) == "shape3dContourRgb" && HasSafeNativePictureShape3dContourRgb(picture) ||
                      LeafKind(operation) == "shape3dContourColorScheme" && HasSafeNativePictureShape3dContourColorScheme(picture) ||
                      LeafKind(operation) == "shape3dExtrusionRgb" && HasSafeNativePictureShape3dExtrusionRgb(picture) ||
                      LeafKind(operation) == "shape3dExtrusionColorScheme" && HasSafeNativePictureShape3dExtrusionColorScheme(picture) ||
                      LeafKind(operation) == "shape3dSceneCameraPreset" && HasSafeNativePictureShape3dSceneCameraPreset(picture) ||
                      LeafKind(operation) == "shape3dSceneCameraZoomThousandthPercent" && HasSafeNativePictureShape3dSceneCameraZoom(picture) ||
                      LeafKind(operation) == "shape3dSceneCameraFov60000" && HasSafeNativePictureShape3dSceneCameraFov(picture) ||
                      LeafKind(operation) == "shape3dSceneCameraRotationLatitude60000" && HasSafeNativePictureShape3dSceneCameraRotationLatitude(picture) ||
                      LeafKind(operation) == "shape3dSceneCameraRotationLongitude60000" && HasSafeNativePictureShape3dSceneCameraRotationLongitude(picture) ||
                      LeafKind(operation) == "shape3dSceneCameraRotationRevolution60000" && HasSafeNativePictureShape3dSceneCameraRotationRevolution(picture) ||
                      LeafKind(operation) == "shape3dSceneLightRigPreset" && HasSafeNativePictureShape3dSceneLightRigPreset(picture) ||
                      LeafKind(operation) == "shape3dSceneLightRigDirection" && HasSafeNativePictureShape3dSceneLightRigDirection(picture) ||
                      LeafKind(operation) == "shape3dSceneLightRigRotationLatitude60000" && HasSafeNativePictureShape3dSceneLightRigRotationLatitude(picture) ||
                      LeafKind(operation) == "shape3dSceneLightRigRotationLongitude60000" && HasSafeNativePictureShape3dSceneLightRigRotationLongitude(picture) ||
                      LeafKind(operation) == "shape3dSceneLightRigRotationRevolution60000" && HasSafeNativePictureShape3dSceneLightRigRotationRevolution(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropAnchorXEmu" && HasSafeNativePictureShape3dSceneBackdropAnchorX(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropAnchorYEmu" && HasSafeNativePictureShape3dSceneBackdropAnchorY(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropAnchorZEmu" && HasSafeNativePictureShape3dSceneBackdropAnchorZ(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropNormalDxEmu" && HasSafeNativePictureShape3dSceneBackdropNormalDx(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropNormalDyEmu" && HasSafeNativePictureShape3dSceneBackdropNormalDy(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropNormalDzEmu" && HasSafeNativePictureShape3dSceneBackdropNormalDz(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropUpDxEmu" && HasSafeNativePictureShape3dSceneBackdropUpDx(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropUpDyEmu" && HasSafeNativePictureShape3dSceneBackdropUpDy(picture) ||
                      LeafKind(operation) == "shape3dSceneBackdropUpDzEmu" && HasSafeNativePictureShape3dSceneBackdropUpDz(picture) ||
                      LeafKind(operation) == "shape3dPresetMaterial" && HasSafeNativePictureShape3dPresetMaterial(picture) ||
                      LeafKind(operation) == "shape3dBevelTopWidthEmu" && HasSafeNativePictureShape3dBevelTopWidth(picture) ||
                      LeafKind(operation) == "shape3dBevelTopHeightEmu" && HasSafeNativePictureShape3dBevelTopHeight(picture) ||
                      LeafKind(operation) == "shape3dBevelTopPreset" && HasSafeNativePictureShape3dBevelTopPreset(picture) ||
                      LeafKind(operation) == "shape3dBevelBottomWidthEmu" && HasSafeNativePictureShape3dBevelBottomWidth(picture) ||
                      LeafKind(operation) == "shape3dBevelBottomHeightEmu" && HasSafeNativePictureShape3dBevelBottomHeight(picture) ||
                      LeafKind(operation) == "shape3dBevelBottomPreset" && HasSafeNativePictureShape3dBevelBottomPreset(picture) ||
                      LeafKind(operation) == "shape3dDepthEmu" && HasSafeNativePictureShape3dDepth(picture) ||
                      LeafKind(operation) is "imageInnerShadowBlurRadiusEmu" or "imageInnerShadowDistanceEmu" or "imageInnerShadowDirectionDegrees" or "imageInnerShadowColorRgb" or "imageInnerShadowColorScheme" or "imageInnerShadowOpacityThousandthPercent" && HasSafeNativePictureInnerShadow(picture, LeafKind(operation)) ||
                      LeafKind(operation) is "imageReflectionBlurRadiusEmu" or "imageReflectionStartOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" or "imageReflectionDistanceEmu" or "imageReflectionDirectionDegrees" && HasSafeNativePictureReflection(picture, LeafKind(operation)) ||
                      LeafKind(operation) == "imageSoftEdgeRadiusEmu" && HasSafeNativePictureSoftEdge(picture)))
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
        if (proofs.All(proof => LeafKind(proof.Operation) is "chartTitleText" || IsChartDataLeafKind(LeafKind(proof.Operation))))
        {
            var chartPatches = new List<PptxXmlPatch>();
            var titleProofs = proofs.Where(proof => LeafKind(proof.Operation) == "chartTitleText").ToArray();
            var dataProofs = proofs.Where(proof => IsChartDataLeafKind(LeafKind(proof.Operation))).ToArray();
            if (titleProofs.Length > 0) chartPatches.AddRange(CompileChartTitleXmlPatches(partBytes, titleProofs));
            if (dataProofs.Length > 0) chartPatches.AddRange(CompileChartDataXmlPatches(partBytes, dataProofs));
            return OrderedNonOverlapping(chartPatches, proofs[0].MutationPartPath);
        }
        if (proofs.Any(proof => LeafKind(proof.Operation) is "chartTitleText" || IsChartDataLeafKind(LeafKind(proof.Operation))))
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
            if (leafKind == "paragraphAlignment")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraphAlignment target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextParagraphAlignmentXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" or "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" or "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextParagraphSpacingXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "paragraphMarginLeftEmu" or "paragraphIndentEmu")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextParagraphLayoutXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "paragraphBulletCharacter")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraphBulletCharacter target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextParagraphBulletCharacterXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "paragraphLevel")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraphLevel target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextParagraphLevelXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "paragraphBulletAutoNumberScheme" or "paragraphBulletAutoNumberStartAt")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextParagraphAutoNumberXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "paragraphBulletFontFamily" or "paragraphBulletColorRgb" or "paragraphBulletColorScheme" or "paragraphBulletSizePoints" or "paragraphBulletSizePercent")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextParagraphBulletStyleXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "verticalAnchor")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} verticalAnchor target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextVerticalAnchorXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textBodyInsetLeftEmu" or "textBodyInsetTopEmu" or "textBodyInsetRightEmu" or "textBodyInsetBottomEmu")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyInsetXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyWrap")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyWrap target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyWrapXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyColumnCount")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyColumnCount target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyColumnCountXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyColumnGapEmu")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyColumnGapEmu target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyColumnGapXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyRotationDegrees")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyRotationDegrees target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyRotationXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyVerticalOverflow")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyVerticalOverflow target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyVerticalOverflowXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyHorizontalOverflow")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyHorizontalOverflow target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyHorizontalOverflowXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyUpright")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyUpright target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyUprightXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyAnchorCenter")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyAnchorCenter target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyAnchorCenterXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyForceAntiAlias")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyForceAntiAlias target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyForceAntiAliasXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodySpaceFirstLastParagraph")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodySpaceFirstLastParagraph target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodySpaceFirstLastParagraphXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyCompatibleLineSpacing")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyCompatibleLineSpacing target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyCompatibleLineSpacingXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyFromWordArt")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyFromWordArt target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyFromWordArtXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyWarpPreset")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyWarpPreset target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyWarpPresetXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyWarpAdjustment")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyWarpAdjustment target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyWarpAdjustmentXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyFlatTextZ")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyFlatTextZ target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyFlatTextZXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyAutoFit")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyAutoFit target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyAutoFitXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textBodyNormalAutoFitFontScale" or "textBodyNormalAutoFitLineSpacingReduction")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyNormalAutoFitXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyColumnDirection")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyColumnDirection target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyColumnDirectionXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textBodyVerticalText")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyVerticalText target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyVerticalTextXmlPatch(xml, range, proof));
                continue;
            }
            if ((leafKind is "rotationDegrees" or "flipHorizontal" or "flipVertical") && range.LocalName is not ("sp" or "pic" or "grpSp"))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
            if (leafKind is "fontSizePoints" or "fontFamily" or "fontFamilyEastAsia" or "fontFamilyComplexScript" or "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontColorRgb" or "fontColorScheme" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps" or "fontHighlightRgb" or "fontHighlightScheme")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(leafKind == "fontSizePoints"
                    ? CompileTextFontSizeXmlPatch(xml, range, proof, drawingPrefixes)
                    : leafKind is "fontFamily" or "fontFamilyEastAsia" or "fontFamilyComplexScript"
                        ? CompileTextFontFamilyXmlPatch(xml, range, proof, drawingPrefixes)
                        : leafKind is "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps"
                            ? CompileTextFontBooleanXmlPatch(xml, range, proof, drawingPrefixes)
                            : CompileTextFontColorXmlPatch(xml, range, proof, drawingPrefixes));
                continue;
            }
            if (leafKind is "textGlowRadiusEmu" or "textGlowColorRgb" or "textGlowColorScheme" or "textGlowOpacityThousandthPercent")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextGlowXmlPatch(xml, range, proof, drawingPrefixes));
                continue;
            }
            if (leafKind is "textDefaultGlowRadiusEmu" or "textDefaultGlowColorRgb" or "textDefaultGlowColorScheme" or "textDefaultGlowOpacityThousandthPercent")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultGlowXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textDefaultShadowBlurRadiusEmu" or "textDefaultShadowDistanceEmu" or "textDefaultShadowDirectionDegrees" or "textDefaultShadowAlignment")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultShadowXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textDefaultShadowColorRgb" or "textDefaultShadowColorScheme")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultShadowColorXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textDefaultShadowOpacityThousandthPercent")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultShadowAlphaXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textDefaultShadowRotateWithShape")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultShadowRotateWithShapeXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textDefaultInnerShadowBlurRadiusEmu" or "textDefaultInnerShadowDistanceEmu" or "textDefaultInnerShadowDirectionDegrees")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textDefaultInnerShadowBlurRadiusEmu target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultInnerShadowXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textDefaultInnerShadowColorRgb" or "textDefaultInnerShadowColorScheme")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultInnerShadowColorXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind == "textDefaultInnerShadowOpacityThousandthPercent")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultInnerShadowAlphaXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textDefaultReflectionBlurRadiusEmu" or "textDefaultReflectionDistanceEmu" or "textDefaultReflectionStartOpacityThousandthPercent" or "textDefaultReflectionEndOpacityThousandthPercent" or "textDefaultReflectionDirectionDegrees")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextDefaultReflectionXmlPatch(xml, range, proof));
                continue;
            }
            if (leafKind is "textInnerShadowBlurRadiusEmu" or "textInnerShadowDistanceEmu" or "textInnerShadowDirectionDegrees" or "textInnerShadowColorRgb" or "textInnerShadowColorScheme" or "textInnerShadowOpacityThousandthPercent")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextInnerShadowXmlPatch(xml, range, proof, drawingPrefixes));
                continue;
            }
            if (leafKind is "textReflectionBlurRadiusEmu" or "textReflectionStartOpacityThousandthPercent" or "textReflectionEndOpacityThousandthPercent" or "textReflectionDistanceEmu" or "textReflectionDirectionDegrees")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextReflectionXmlPatch(xml, range, proof, drawingPrefixes));
                continue;
            }
            if (leafKind is "textSoftEdgeRadiusEmu" or "textDefaultSoftEdgeRadiusEmu")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(leafKind == "textSoftEdgeRadiusEmu"
                    ? CompileTextSoftEdgeXmlPatch(xml, range, proof, drawingPrefixes)
                    : CompileTextDefaultSoftEdgeXmlPatch(xml, range, proof));
                continue;
            }
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
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (leaves.Length == 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} table cell has no bounded text leaves.", operation.SlidePartPath);
        var sourceLines = leaves.Select(leaf => DecodeTextLeaf(leaf.Value, leaf.Groups["prefix"].Value)).ToArray();
        if (string.Join("\n", sourceLines) != operation.ExpectedValue)
            throw new CodecException("presentation_text_precondition_failed", $"PPTX edit operation {operation.OperationId} raw table-cell leaves do not match the expected text.", operation.SlidePartPath);
        var requestedLines = operation.Value.Split('\n');
        if (requestedLines.Length != leaves.Length)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} must preserve the source paragraph count when editing a multi-paragraph table cell.", operation.SlidePartPath);

        // Replace all paragraph leaves inside one source span so the wire
        // result remains one logical operation/result while preserving every
        // paragraph, run property and interstitial XML token byte-for-byte.
        var firstStart = leaves[0].Index;
        var lastEnd = leaves[^1].Index + leaves[^1].Length;
        var replacement = cell.Value[firstStart..lastEnd];
        for (var index = leaves.Length - 1; index >= 0; index--)
        {
            var leaf = leaves[index];
            var open = leaf.Groups["open"].Value;
            if (NeedsPreserve(requestedLines[index]) && !PreserveSpacePattern().IsMatch(open))
                open = open.Insert(open.Length - 1, " xml:space=\"preserve\"");
            var replacementLeaf = ReplaceTextLeaf(leaf, leaf.Groups["prefix"].Value, requestedLines[index], open);
            var start = leaf.Index - firstStart;
            var end = start + leaf.Length;
            replacement = replacement[..start] + replacementLeaf + replacement[end..];
        }
        var absoluteStart = elementRange.Start + cell.Index + firstStart;
        return new PptxXmlPatch(
            operation,
            absoluteStart,
            absoluteStart + (lastEnd - firstStart),
            replacement,
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
        if (LeafKind(operation) == "customGeometryPathWidth" &&
            (element is not P.Shape pathWidthShape || !TryRequestedNativeCustomGeometryPathWidth(pathWidthShape, operation.NativeLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry path width must be a changed positive canonical coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathHeight" &&
            (element is not P.Shape pathHeightShape || !TryRequestedNativeCustomGeometryPathHeight(pathHeightShape, operation.NativeLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry path height must be a changed positive canonical coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathArcWidthRadius" &&
            (element is not P.Shape pathArcWidthRadiusShape || !TryRequestedNativeCustomGeometryPathArcWidthRadius(pathArcWidthRadiusShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry arc width radius must be a changed positive canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathArcHeightRadius" &&
            (element is not P.Shape pathArcHeightRadiusShape || !TryRequestedNativeCustomGeometryPathArcHeightRadius(pathArcHeightRadiusShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry arc height radius must be a changed positive canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathArcStartAngle60000" &&
            (element is not P.Shape pathArcStartAngleShape || !TryRequestedNativeCustomGeometryPathArcStartAngle(pathArcStartAngleShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry arc start angle must be a changed canonical one-turn angle.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathArcSweepAngle60000" &&
            (element is not P.Shape pathArcSweepAngleShape || !TryRequestedNativeCustomGeometryPathArcSweepAngle(pathArcSweepAngleShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry arc sweep angle must be a changed non-zero canonical one-turn angle.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dExtrusionHeightEmu" &&
            !TryRequestedShape3dExtrusionHeight(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D extrusion height must be a changed canonical non-negative coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dDepthEmu" &&
            !TryRequestedShape3dDepth(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D depth must be a changed canonical signed coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dContourWidthEmu" &&
            !TryRequestedShape3dContourWidth(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D contour width must be a changed canonical non-negative coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dContourRgb" &&
            !TryRequestedShape3dContourRgb(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape 3-D contour color must be a changed canonical RGB token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dContourColorScheme" &&
            !TryRequestedShape3dContourColorScheme(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D contour theme color must be a changed canonical theme token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dExtrusionRgb" &&
            !TryRequestedShape3dExtrusionRgb(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape 3-D extrusion color must be a changed canonical RGB token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dExtrusionColorScheme" &&
            !TryRequestedShape3dExtrusionColorScheme(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D extrusion theme color must be a changed canonical theme token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneCameraPreset" &&
            !TryRequestedShape3dSceneCameraPreset(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D scene camera preset must be a changed canonical camera-preset token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneCameraZoomThousandthPercent" &&
            !TryRequestedShape3dSceneCameraZoom(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D scene camera zoom must be a changed canonical non-negative zoom token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneCameraFov60000" &&
            !TryRequestedShape3dSceneCameraFov(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D scene camera FOV must be a changed canonical positive FOV token below 180 degrees.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneCameraRotationLatitude60000" &&
            !TryRequestedShape3dSceneCameraRotationLatitude(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D camera rotation latitude must be a changed canonical non-negative angle token at or below 360 degrees.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneCameraRotationLongitude60000" &&
            !TryRequestedShape3dSceneCameraRotationLongitude(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D camera rotation longitude must be a changed canonical non-negative angle token at or below 360 degrees.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneCameraRotationRevolution60000" &&
            !TryRequestedShape3dSceneCameraRotationRevolution(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D camera rotation revolution must be a changed canonical non-negative angle token at or below 360 degrees.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropAnchorXEmu" &&
            !TryRequestedShape3dSceneBackdropAnchorX(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop anchor X must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropAnchorYEmu" &&
            !TryRequestedShape3dSceneBackdropAnchorY(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop anchor Y must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropAnchorZEmu" &&
            !TryRequestedShape3dSceneBackdropAnchorZ(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop anchor Z must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropNormalDxEmu" &&
            !TryRequestedShape3dSceneBackdropNormalDx(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop normal X must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropNormalDyEmu" &&
            !TryRequestedShape3dSceneBackdropNormalDy(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop normal Y must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropNormalDzEmu" &&
            !TryRequestedShape3dSceneBackdropNormalDz(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop normal Z must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropUpDxEmu" &&
            !TryRequestedShape3dSceneBackdropUpDx(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop up-vector X must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropUpDyEmu" &&
            !TryRequestedShape3dSceneBackdropUpDy(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop up-vector Y must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneBackdropUpDzEmu" &&
            !TryRequestedShape3dSceneBackdropUpDz(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D backdrop up-vector Z must be a changed canonical signed EMU coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneLightRigPreset" &&
            !TryRequestedShape3dSceneLightRigPreset(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D scene light-rig preset must be a changed canonical light-rig preset token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneLightRigDirection" &&
            !TryRequestedShape3dSceneLightRigDirection(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D scene light-rig direction must be a changed canonical light-rig direction token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneLightRigRotationLatitude60000" &&
            !TryRequestedShape3dSceneLightRigRotationLatitude(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D light-rig rotation latitude must be a changed canonical non-negative angle token at or below 360 degrees.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneLightRigRotationLongitude60000" &&
            !TryRequestedShape3dSceneLightRigRotationLongitude(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape 3-D light-rig rotation longitude must be a changed canonical non-negative angle token at or below 360 degrees.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dSceneLightRigRotationRevolution60000" &&
            !TryRequestedShape3dSceneLightRigRotationRevolution(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D light-rig rotation revolution must be a changed canonical non-negative angle token at or below 360 degrees.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dPresetMaterial" &&
            !TryRequestedShape3dPresetMaterial(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D material must be a changed canonical preset-material token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dBevelTopWidthEmu" &&
            !TryRequestedShape3dBevelTopWidth(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D top-bevel width must be a changed canonical non-negative coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dBevelTopHeightEmu" &&
            !TryRequestedShape3dBevelTopHeight(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D top-bevel height must be a changed canonical non-negative coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dBevelTopPreset" &&
            !TryRequestedShape3dBevelTopPreset(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D top-bevel preset must be a changed canonical preset token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dBevelBottomWidthEmu" &&
            !TryRequestedShape3dBevelBottomWidth(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D bottom-bevel width must be a changed canonical non-negative coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dBevelBottomHeightEmu" &&
            !TryRequestedShape3dBevelBottomHeight(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D bottom-bevel height must be a changed canonical non-negative coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "shape3dBevelBottomPreset" &&
            !TryRequestedShape3dBevelBottomPreset(element, operation.Value, out _))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} shape or picture 3-D bottom-bevel preset must be a changed canonical preset token.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathLineToX" &&
            (element is not P.Shape pathLineToXShape || !TryRequestedNativeCustomGeometryPathLineToX(pathLineToXShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry line-to x coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathLineToY" &&
            (element is not P.Shape pathLineToYShape || !TryRequestedNativeCustomGeometryPathLineToY(pathLineToYShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry line-to y coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathMoveToX" &&
            (element is not P.Shape pathMoveToXShape || !TryRequestedNativeCustomGeometryPathMoveToX(pathMoveToXShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry move-to x coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathMoveToY" &&
            (element is not P.Shape pathMoveToYShape || !TryRequestedNativeCustomGeometryPathMoveToY(pathMoveToYShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry move-to y coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathQuadraticEndX" &&
            (element is not P.Shape pathQuadraticEndXShape || !TryRequestedNativeCustomGeometryPathQuadraticEndX(pathQuadraticEndXShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry quadratic end-point x coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathQuadraticEndY" &&
            (element is not P.Shape pathQuadraticEndYShape || !TryRequestedNativeCustomGeometryPathQuadraticEndY(pathQuadraticEndYShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry quadratic end-point y coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathQuadraticControlX" &&
            (element is not P.Shape pathQuadraticControlXShape || !TryRequestedNativeCustomGeometryPathQuadraticControlX(pathQuadraticControlXShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry quadratic control-point x coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathQuadraticControlY" &&
            (element is not P.Shape pathQuadraticControlYShape || !TryRequestedNativeCustomGeometryPathQuadraticControlY(pathQuadraticControlYShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry quadratic control-point y coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathCubicEndX" &&
            (element is not P.Shape pathCubicEndXShape || !TryRequestedNativeCustomGeometryPathCubicEndX(pathCubicEndXShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry cubic end-point x coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathCubicEndY" &&
            (element is not P.Shape pathCubicEndYShape || !TryRequestedNativeCustomGeometryPathCubicEndY(pathCubicEndYShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry cubic end-point y coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathCubicControl1X" &&
            (element is not P.Shape pathCubicControl1XShape || !TryRequestedNativeCustomGeometryPathCubicControl1X(pathCubicControl1XShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry cubic first-control-point x coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathCubicControl1Y" &&
            (element is not P.Shape pathCubicControl1YShape || !TryRequestedNativeCustomGeometryPathCubicControl1Y(pathCubicControl1YShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry cubic first-control-point y coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathCubicControl2X" &&
            (element is not P.Shape pathCubicControl2XShape || !TryRequestedNativeCustomGeometryPathCubicControl2X(pathCubicControl2XShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry cubic second-control-point x coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathCubicControl2Y" &&
            (element is not P.Shape pathCubicControl2YShape || !TryRequestedNativeCustomGeometryPathCubicControl2Y(pathCubicControl2YShape, operation.NativeLeafIndex, operation.TextLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry cubic second-control-point y coordinate must be a changed canonical path coordinate.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathFill" &&
            (element is not P.Shape pathFillShape || !TryRequestedNativeCustomGeometryPathFill(pathFillShape, operation.NativeLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry path fill must be a changed canonical boolean.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathStroke" &&
            (element is not P.Shape pathStrokeShape || !TryRequestedNativeCustomGeometryPathStroke(pathStrokeShape, operation.NativeLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry path stroke must be a changed canonical boolean.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryPathExtrusionAllowed" &&
            (element is not P.Shape pathExtrusionShape || !TryRequestedNativeCustomGeometryPathExtrusionAllowed(pathExtrusionShape, operation.NativeLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry path extrusion permission must be a changed canonical boolean.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryAdjustmentFormula" &&
            (element is not P.Shape adjustmentFormulaShape || !TryRequestedNativeCustomGeometryAdjustmentFormula(adjustmentFormulaShape, operation.NativeLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry adjustment formula must be a changed canonical formula accepted by the existing guide graph.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryGuideFormula" &&
            (element is not P.Shape guideFormulaShape || !TryRequestedNativeCustomGeometryGuideFormula(guideFormulaShape, operation.NativeLeafIndex, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} custom-geometry guide formula must be a changed canonical formula accepted by the existing guide graph.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryTextRectangleLeftEmu" &&
            (element is not P.Shape leftShape || !TryRequestedNativeCustomGeometryTextRectangleLeft(leftShape, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} text-rectangle left edge must be a changed canonical in-frame coordinate smaller than the right edge.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryTextRectangleTopEmu" &&
            (element is not P.Shape topShape || !TryRequestedNativeCustomGeometryTextRectangleTop(topShape, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} text-rectangle top edge must be a changed canonical in-frame coordinate smaller than the bottom edge.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryTextRectangleRightEmu" &&
            (element is not P.Shape rightShape || !TryRequestedNativeCustomGeometryTextRectangleRight(rightShape, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} text-rectangle right edge must be a changed canonical in-frame coordinate greater than the left edge.",
                operation.SlidePartPath);
        if (LeafKind(operation) == "customGeometryTextRectangleBottomEmu" &&
            (element is not P.Shape bottomShape || !TryRequestedNativeCustomGeometryTextRectangleBottom(bottomShape, operation.Value)))
            throw new CodecException(
                "invalid_presentation_edit_operation",
                $"PPTX edit operation {operation.OperationId} text-rectangle bottom edge must be a changed canonical in-frame coordinate greater than the top edge.",
                operation.SlidePartPath);
    }

    private static bool HasSafeNativeConnectorLine(P.ConnectionShape connector, string kind)
    {
        var outline = connector.ShapeProperties?.Elements<A.Outline>().ToArray();
        if (outline is not { Length: 1 }) return false;
        if (kind == "lineWidthEmu")
            return outline[0].Width?.Value is { } width && width is >= 0 and <= 20_116_800;
        if (kind == "lineOpacityThousandthPercent")
        {
            var fills = outline[0].Elements<A.SolidFill>().ToArray();
            return fills.Length == 1 &&
                (HasSafeNativeRgbFillOpacity(fills[0]) || HasSafeNativeSchemeFillOpacity(fills[0]));
        }
        if (kind is "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength")
        {
            var fills = outline[0].Elements<A.SolidFill>().ToArray();
            var hasExplicitPaint = fills.Length == 1 &&
                (HasSafeNativeRgbFill(fills[0]) || HasSafeNativeSchemeFill(fills[0]));
            if (kind == "lineStyle")
            {
                if (!hasExplicitPaint) return false;
                var dashes = outline[0].Elements<A.PresetDash>().ToArray();
                return dashes.Length == 1 && PptxLineStyleCodec.TryReadPresetDash(dashes[0], out _);
            }
            if (kind == "lineCap") return hasExplicitPaint && PptxLineStyleCodec.TryReadCap(outline[0], out _);
            if (kind is "lineStartArrow" or "lineEndArrow")
            {
                OpenXmlElement? endpoint = kind == "lineStartArrow"
                    ? outline[0].GetFirstChild<A.HeadEnd>()
                    : outline[0].GetFirstChild<A.TailEnd>();
                return PptxLineStyleCodec.TryReadArrowType(endpoint, out _) &&
                    (hasExplicitPaint || HasSafeNativeConnectorStyleReference(connector));
            }
            if (kind is "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength")
            {
                OpenXmlElement? endpoint = kind is "lineStartArrowWidth" or "lineStartArrowLength"
                    ? outline[0].GetFirstChild<A.HeadEnd>()
                    : outline[0].GetFirstChild<A.TailEnd>();
                return PptxLineStyleCodec.TryReadArrowSize(endpoint,
                           kind is "lineStartArrowWidth" or "lineEndArrowWidth", out _) &&
                    (hasExplicitPaint || HasSafeNativeConnectorStyleReference(connector));
            }
            return hasExplicitPaint && PptxLineStyleCodec.TryReadJoinLeaf(outline[0], out _);
        }
        var solidFill = outline[0].Elements<A.SolidFill>().ToArray();
        if (solidFill.Length == 1 && solidFill[0].ChildElements.Count == 1)
        {
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
        }
        if (solidFill.Length != 0 || kind is not ("lineRgb" or "lineScheme")) return false;
        return TryReadNativeConnectorStyleColor(connector, kind, out _);
    }

    private static bool TryReadNativeConnectorStyleColor(P.ConnectionShape connector, string kind, out string value)
    {
        value = string.Empty;
        if (!HasSafeNativeConnectorStyleReference(connector)) return false;
        var color = connector.ShapeStyle?.LineReference?.FirstChild;
        if (kind == "lineScheme" && color is A.SchemeColor scheme && scheme.Val?.Value is { } schemeValue &&
            PptxColor.TrySchemeToken(schemeValue, out var schemeToken))
        {
            value = schemeToken;
            return true;
        }
        if (kind == "lineRgb" && color is A.RgbColorModelHex rgb && rgb.Val?.Value is { Length: 6 } rgbValue &&
            rgbValue.All(Uri.IsHexDigit))
        {
            value = rgbValue.ToUpperInvariant();
            return true;
        }
        return false;
    }

    private static bool HasSafeNativeConnectorStyleReference(P.ConnectionShape connector)
    {
        var style = connector.ShapeStyle;
        if (style is null || style.GetAttributes().Count != 0 || style.ChildElements.Count != 4 ||
            style.ChildElements.Count(child => child is A.LineReference) != 1 ||
            style.ChildElements.Count(child => child is A.FillReference) != 1 ||
            style.ChildElements.Count(child => child is A.EffectReference) != 1 ||
            style.ChildElements.Count(child => child is A.FontReference) != 1)
            return false;
        return HasSafeNativeStyleReference(style.LineReference, font: false) &&
            HasSafeNativeStyleReference(style.FillReference, font: false) &&
            HasSafeNativeStyleReference(style.EffectReference, font: false) &&
            HasSafeNativeStyleReference(style.FontReference, font: true);
    }

    private static bool HasSafeNativeStyleReference(OpenXmlElement? reference, bool font)
    {
        if (reference is null || reference.ChildElements.Count != 1) return false;
        var attributes = reference.GetAttributes();
        if (attributes.Count != 1 || attributes[0].LocalName != "idx" || attributes[0].NamespaceUri.Length != 0)
            return false;
        var index = attributes[0].Value;
        if (font)
        {
            if (index is not ("minor" or "major")) return false;
        }
        else if (!uint.TryParse(index, out var numericIndex) || numericIndex > 32)
        {
            return false;
        }
        var color = reference.FirstChild;
        if (color is A.SchemeColor scheme)
            return scheme.ChildElements.Count == 0 &&
                scheme.GetAttributes().Count == 1 && scheme.GetAttributes()[0].LocalName == "val" &&
                scheme.GetAttributes()[0].NamespaceUri.Length == 0 && scheme.Val?.Value is { } schemeValue &&
                PptxColor.TrySchemeToken(schemeValue, out _);
        if (color is A.RgbColorModelHex rgb)
            return rgb.ChildElements.Count == 0 &&
                rgb.GetAttributes().Count == 1 && rgb.GetAttributes()[0].LocalName == "val" &&
                rgb.GetAttributes()[0].NamespaceUri.Length == 0 && rgb.Val?.Value is { Length: 6 } rgbValue &&
                rgbValue.All(Uri.IsHexDigit);
        return false;
    }

    private static bool HasSafeNativeShapeStyle(P.Shape shape, string kind)
    {
        var properties = shape.ShapeProperties;
        if (properties is null) return false;
        if (kind is "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment" or "shadowRotateWithShape" or "imageShadowBlurRadiusEmu" or "imageShadowDistanceEmu" or "imageShadowDirectionDegrees" or "imageShadowAlignment")
            return HasSafeNativeShadowGeometry(properties, kind);
        if (kind == "shadowOpacityThousandthPercent")
            return HasSafeNativeShadowOpacity(properties);
        if (kind is "shadowColorRgb" or "shadowColorScheme")
            return HasSafeNativeShadowColor(properties, kind);
        if (kind is "fillRgb" or "fillOpacityThousandthPercent" or "fillScheme")
        {
            var fills = properties.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray();
            return fills.Length == 1 && fills[0] is A.SolidFill solid &&
                (kind == "fillRgb" ? HasSafeNativeRgbFill(solid) :
                 kind == "fillOpacityThousandthPercent" ? HasSafeNativeRgbFillOpacity(solid) || HasSafeNativeSchemeFillOpacity(solid) :
                 HasSafeNativeSchemeFill(solid));
        }
        var outlines = properties.Elements<A.Outline>().ToArray();
        if (outlines.Length != 1) return false;
        if (kind == "lineWidthEmu")
            return outlines[0].Width?.Value is > 0 and <= 20_116_800;
        var linePaints = outlines[0].ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray();
        if (kind == "lineOpacityThousandthPercent")
            return linePaints.Length == 1 && linePaints[0] is A.SolidFill lineOpacitySolid &&
                (HasSafeNativeRgbFillOpacity(lineOpacitySolid) || HasSafeNativeSchemeFillOpacity(lineOpacitySolid));
        var hasExplicitPaint = linePaints.Length == 1 && linePaints[0] is A.SolidFill lineSolid &&
            (HasSafeNativeRgbFill(lineSolid) || HasSafeNativeSchemeFill(lineSolid));
        if (kind == "lineStyle")
        {
            var dashes = outlines[0].Elements<A.PresetDash>().ToArray();
            return hasExplicitPaint && dashes.Length == 1 && PptxLineStyleCodec.TryReadPresetDash(dashes[0], out _);
        }
        if (kind == "lineCap") return hasExplicitPaint && PptxLineStyleCodec.TryReadCap(outlines[0], out _);
        if (kind == "lineJoin") return hasExplicitPaint && PptxLineStyleCodec.TryReadJoinLeaf(outlines[0], out _);
        if (kind is "lineStartArrow" or "lineEndArrow")
        {
            OpenXmlElement? endpoint = kind == "lineStartArrow"
                ? outlines[0].GetFirstChild<A.HeadEnd>()
                : outlines[0].GetFirstChild<A.TailEnd>();
            return hasExplicitPaint && PptxLineStyleCodec.TryReadArrowType(endpoint, out var arrow) && arrow.Length > 0;
        }
        if (kind is "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength")
        {
            OpenXmlElement? endpoint = kind is "lineStartArrowWidth" or "lineStartArrowLength"
                ? outlines[0].GetFirstChild<A.HeadEnd>()
                : outlines[0].GetFirstChild<A.TailEnd>();
            return hasExplicitPaint && PptxLineStyleCodec.TryReadArrowSize(endpoint,
                kind is "lineStartArrowWidth" or "lineEndArrowWidth", out _);
        }
        if (kind is not ("lineRgb" or "lineScheme")) return false;
        return linePaints.Length == 1 && linePaints[0] is A.SolidFill directLineSolid &&
            (kind == "lineRgb" && HasSafeNativeRgbFill(directLineSolid) ||
             kind == "lineScheme" && HasSafeNativeSchemeFill(directLineSolid));
    }

    private static bool HasSafeNativeTextWarpAdjustment(P.Shape shape, uint nativeIndex)
    {
        var bodyProperties = shape.TextBody?.BodyProperties;
        if (bodyProperties is null || !PptxBodyPropertiesCodec.TryReadTextWarp(bodyProperties, out _, out var adjustments))
            return false;
        return nativeIndex < (uint)adjustments.Count;
    }

    private static bool HasSafeNativeCustomGeometryAdjustment(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var adjustments = geometry?.GetFirstChild<A.AdjustValueList>();
        if (geometry is null || adjustments is null || geometry.HasAttributes || adjustments.HasAttributes ||
            adjustments.ChildElements.Count != adjustments.Elements<A.ShapeGuide>().Count())
            return false;
        var guides = adjustments.Elements<A.ShapeGuide>().ToArray();
        if (nativeIndex >= (uint)guides.Length) return false;
        var guide = guides[nativeIndex];
        if (guide.ChildElements.Count != 0 || guide.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")) ||
            guide.Name?.Value is not { Length: > 0 } || guide.Formula?.Value is not { } formula)
            return false;
        return TryLiteralCustomAdjustment(formula, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentFormula(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var adjustments = geometry?.GetFirstChild<A.AdjustValueList>();
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (geometry is null || adjustments is null || shapeWidth is null || shapeHeight is null ||
            shapeWidth <= 0 || shapeHeight <= 0 || geometry.HasAttributes || adjustments.HasAttributes ||
            geometry.Elements<A.AdjustValueList>().Count() != 1 ||
            adjustments.ChildElements.Count != adjustments.Elements<A.ShapeGuide>().Count())
            return false;
        var guides = adjustments.Elements<A.ShapeGuide>().ToArray();
        if (nativeIndex >= (uint)guides.Length) return false;
        foreach (var guide in guides)
        {
            if (guide.ChildElements.Count != 0 || guide.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")) ||
                guide.Name?.Value is not { Length: > 0 } ||
                guide.Formula?.Value is not { Length: > 0 })
                return false;
        }
        var sourceFormula = guides[(int)nativeIndex].Formula?.Value;
        if (sourceFormula is not { Length: > 0 }) return false;
        return !TryLiteralCustomAdjustment(sourceFormula, out _) &&
            PptxCustomGeometryCodec.TryReadCanonicalAdjustmentFormula(
                geometry, shapeWidth.Value, shapeHeight.Value, nativeIndex, out var canonicalFormula) &&
            sourceFormula == canonicalFormula;
    }

    private static bool TryRequestedNativeCustomGeometryAdjustmentFormula(P.Shape shape, uint nativeIndex, string token)
    {
        if (!HasSafeNativeCustomGeometryAdjustmentFormula(shape, nativeIndex)) return false;
        var geometry = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!;
        var guides = geometry.GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>().ToArray();
        var current = guides[(int)nativeIndex].Formula!.Value;
        if (token.Length == 0 || token == current || TryLiteralCustomAdjustment(token, out _))
            return false;
        var shapeWidth = shape.ShapeProperties.Transform2D?.Extents?.Cx?.Value;
        var shapeHeight = shape.ShapeProperties.Transform2D?.Extents?.Cy?.Value;
        if (shapeWidth is null || shapeHeight is null ||
            geometry.CloneNode(true) is not A.CustomGeometry candidate)
            return false;
        var candidateGuides = candidate.GetFirstChild<A.AdjustValueList>()?.Elements<A.ShapeGuide>().ToArray();
        if (candidateGuides is null || nativeIndex >= (uint)candidateGuides.Length)
            return false;
        candidateGuides[(int)nativeIndex].Formula = token;
        return PptxCustomGeometryCodec.TryReadCanonicalAdjustmentFormula(
                candidate, shapeWidth.Value, shapeHeight.Value, nativeIndex, out var canonicalFormula) &&
            canonicalFormula == token;
    }

    private static bool HasSafeNativeCustomGeometryGuide(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var guideList = geometry?.GetFirstChild<A.ShapeGuideList>();
        if (geometry is null || guideList is null || geometry.HasAttributes || guideList.HasAttributes ||
            geometry.Elements<A.ShapeGuideList>().Count() != 1 ||
            guideList.ChildElements.Count != guideList.Elements<A.ShapeGuide>().Count())
            return false;
        var guides = guideList.Elements<A.ShapeGuide>().ToArray();
        if (nativeIndex >= (uint)guides.Length) return false;
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var guide in guides)
        {
            if (guide.ChildElements.Count != 0 || guide.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")) ||
                guide.Name?.Value is not { Length: > 0 } name || !names.Add(name) ||
                guide.Formula?.Value is not { Length: > 0 })
                return false;
        }
        var formula = guides[(int)nativeIndex].Formula?.Value;
        return formula is not null && TryLiteralCustomAdjustment(formula, out _);
    }

    private static bool HasSafeNativeCustomGeometryGuideFormula(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var guideList = geometry?.GetFirstChild<A.ShapeGuideList>();
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (geometry is null || guideList is null || shapeWidth is null || shapeHeight is null ||
            shapeWidth <= 0 || shapeHeight <= 0 || geometry.HasAttributes || guideList.HasAttributes ||
            geometry.Elements<A.ShapeGuideList>().Count() != 1 ||
            guideList.ChildElements.Count != guideList.Elements<A.ShapeGuide>().Count())
            return false;
        var guides = guideList.Elements<A.ShapeGuide>().ToArray();
        if (nativeIndex >= (uint)guides.Length) return false;
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var guide in guides)
        {
            if (guide.ChildElements.Count != 0 || guide.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")) ||
                guide.Name?.Value is not { Length: > 0 } name || !names.Add(name) ||
                guide.Formula?.Value is not { Length: > 0 })
                return false;
        }
        var sourceFormula = guides[(int)nativeIndex].Formula?.Value;
        if (sourceFormula is not { Length: > 0 }) return false;
        return !TryLiteralCustomAdjustment(sourceFormula, out _) &&
            PptxCustomGeometryCodec.TryReadCanonicalGuideFormula(
                geometry, shapeWidth.Value, shapeHeight.Value, nativeIndex, out var canonicalFormula) &&
            sourceFormula == canonicalFormula;
    }

    private static bool TryRequestedNativeCustomGeometryGuideFormula(P.Shape shape, uint nativeIndex, string token)
    {
        if (!HasSafeNativeCustomGeometryGuideFormula(shape, nativeIndex)) return false;
        var geometry = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!;
        var guides = geometry.GetFirstChild<A.ShapeGuideList>()!.Elements<A.ShapeGuide>().ToArray();
        var current = guides[(int)nativeIndex].Formula!.Value;
        if (token.Length == 0 || token == current || TryLiteralCustomAdjustment(token, out _))
            return false;
        var shapeWidth = shape.ShapeProperties.Transform2D?.Extents?.Cx?.Value;
        var shapeHeight = shape.ShapeProperties.Transform2D?.Extents?.Cy?.Value;
        if (shapeWidth is null || shapeHeight is null ||
            geometry.CloneNode(true) is not A.CustomGeometry candidate)
            return false;
        var candidateGuides = candidate.GetFirstChild<A.ShapeGuideList>()?.Elements<A.ShapeGuide>().ToArray();
        if (candidateGuides is null || nativeIndex >= (uint)candidateGuides.Length)
            return false;
        candidateGuides[(int)nativeIndex].Formula = token;
        return PptxCustomGeometryCodec.TryReadCanonicalGuideFormula(
                candidate, shapeWidth.Value, shapeHeight.Value, nativeIndex, out var canonicalFormula) &&
            canonicalFormula == token;
    }

    private static bool HasSafeNativeCustomGeometryConnectionSiteAngle(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var connectionList = geometry?.GetFirstChild<A.ConnectionSiteList>();
        if (geometry is null || connectionList is null || geometry.HasAttributes || connectionList.HasAttributes ||
            geometry.Elements<A.ConnectionSiteList>().Count() != 1 ||
            connectionList.ChildElements.Count != connectionList.Elements<A.ConnectionSite>().Count())
            return false;
        var sites = connectionList.Elements<A.ConnectionSite>().ToArray();
        if (nativeIndex >= (uint)sites.Length) return false;
        foreach (var site in sites)
        {
            if (site.ChildElements.Count != 1 || site.ChildElements[0] is not A.Position position ||
                site.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName != "ang") ||
                position.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                position.ChildElements.Count != 0 || site.Angle?.Value is not { Length: > 0 } ||
                position.X?.Value is not { Length: > 0 } || position.Y?.Value is not { Length: > 0 })
                return false;
        }
        var angle = sites[(int)nativeIndex].Angle?.Value;
        return angle is not null && TryLiteralConnectionSiteAngle(angle, out _);
    }

    private static bool HasSafeNativeCustomGeometryConnectionSiteX(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var connectionList = geometry?.GetFirstChild<A.ConnectionSiteList>();
        if (geometry is null || connectionList is null || geometry.HasAttributes || connectionList.HasAttributes ||
            geometry.Elements<A.ConnectionSiteList>().Count() != 1 ||
            connectionList.ChildElements.Count != connectionList.Elements<A.ConnectionSite>().Count())
            return false;
        var sites = connectionList.Elements<A.ConnectionSite>().ToArray();
        if (nativeIndex >= (uint)sites.Length) return false;
        foreach (var site in sites)
        {
            if (site.ChildElements.Count != 1 || site.ChildElements[0] is not A.Position position ||
                site.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName != "ang") ||
                position.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                position.ChildElements.Count != 0 || site.Angle?.Value is not { Length: > 0 } ||
                position.X?.Value is not { Length: > 0 } || position.Y?.Value is not { Length: > 0 })
                return false;
        }
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        if (shapeWidth is null) return false;
        var selectedPosition = (A.Position)sites[(int)nativeIndex].ChildElements[0];
        var x = selectedPosition.X?.Value;
        return x is not null && TryLiteralConnectionSiteX(x, shapeWidth.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryConnectionSiteY(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var connectionList = geometry?.GetFirstChild<A.ConnectionSiteList>();
        if (geometry is null || connectionList is null || geometry.HasAttributes || connectionList.HasAttributes ||
            geometry.Elements<A.ConnectionSiteList>().Count() != 1 ||
            connectionList.ChildElements.Count != connectionList.Elements<A.ConnectionSite>().Count())
            return false;
        var sites = connectionList.Elements<A.ConnectionSite>().ToArray();
        if (nativeIndex >= (uint)sites.Length) return false;
        foreach (var site in sites)
        {
            if (site.ChildElements.Count != 1 || site.ChildElements[0] is not A.Position position ||
                site.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName != "ang") ||
                position.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                position.ChildElements.Count != 0 || site.Angle?.Value is not { Length: > 0 } ||
                position.X?.Value is not { Length: > 0 } || position.Y?.Value is not { Length: > 0 })
                return false;
        }
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (shapeHeight is null) return false;
        var selectedPosition = (A.Position)sites[(int)nativeIndex].ChildElements[0];
        var y = selectedPosition.Y?.Value;
        return y is not null && TryLiteralConnectionSiteY(y, shapeHeight.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandleX(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandleXY selectedHandle)
            return false;
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        if (shapeWidth is null) return false;
        var position = selectedHandle.GetFirstChild<A.Position>();
        var x = position?.X?.Value;
        return x is not null && TryLiteralCustomGeometryHandleX(x, shapeWidth.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandleY(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandleXY selectedHandle)
            return false;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (shapeHeight is null) return false;
        var position = selectedHandle.GetFirstChild<A.Position>();
        var y = position?.Y?.Value;
        return y is not null && TryLiteralCustomGeometryHandleY(y, shapeHeight.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandleMinX(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandleXY selectedHandle)
            return false;
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        if (shapeWidth is null) return false;
        var minimum = selectedHandle.MinX?.Value;
        return minimum is not null && TryLiteralCustomGeometryHandleMinX(minimum, shapeWidth.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandleMaxX(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandleXY selectedHandle)
            return false;
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        if (shapeWidth is null) return false;
        var maximum = selectedHandle.MaxX?.Value;
        return maximum is not null && TryLiteralCustomGeometryHandleMaxX(maximum, shapeWidth.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandleMinY(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandleXY selectedHandle)
            return false;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (shapeHeight is null) return false;
        var minimum = selectedHandle.MinY?.Value;
        return minimum is not null && TryLiteralCustomGeometryHandleMinY(minimum, shapeHeight.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandleMaxY(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandleXY selectedHandle)
            return false;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (shapeHeight is null) return false;
        var maximum = selectedHandle.MaxY?.Value;
        return maximum is not null && TryLiteralCustomGeometryHandleMaxY(maximum, shapeHeight.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandlePolarMinRadius(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandlePolar selectedHandle)
            return false;
        var minimum = selectedHandle.MinRadial?.Value;
        return minimum is not null && TryLiteralCustomGeometryHandlePolarMinRadius(minimum, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandlePolarMaxRadius(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandlePolar selectedHandle)
            return false;
        var maximum = selectedHandle.MaxRadial?.Value;
        return maximum is not null && TryLiteralCustomGeometryHandlePolarMaxRadius(maximum, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandlePolarMinAngle(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandlePolar selectedHandle)
            return false;
        var minimum = selectedHandle.MinAngle?.Value;
        return minimum is not null && TryLiteralCustomGeometryHandlePolarMinAngle(minimum, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandlePolarMaxAngle(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandlePolar selectedHandle)
            return false;
        var maximum = selectedHandle.MaxAngle?.Value;
        return maximum is not null && TryLiteralCustomGeometryHandlePolarMaxAngle(maximum, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandlePolarX(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandlePolar selectedHandle)
            return false;
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        if (shapeWidth is null) return false;
        var position = selectedHandle.GetFirstChild<A.Position>();
        var x = position?.X?.Value;
        return x is not null && TryLiteralCustomGeometryHandleX(x, shapeWidth.Value, out _);
    }

    private static bool HasSafeNativeCustomGeometryAdjustmentHandlePolarY(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var handleList = geometry?.GetFirstChild<A.AdjustHandleList>();
        if (geometry is null || handleList is null || geometry.HasAttributes || handleList.HasAttributes ||
            geometry.Elements<A.AdjustHandleList>().Count() != 1 ||
            handleList.ChildElements.Count != handleList.Elements<A.AdjustHandleXY>().Count() + handleList.Elements<A.AdjustHandlePolar>().Count())
            return false;
        var handles = handleList.ChildElements.ToArray();
        if (nativeIndex >= (uint)handles.Length) return false;
        foreach (var handle in handles)
        {
            if (handle is A.AdjustHandleXY xy)
            {
                if (xy.ChildElements.Count != 1 || xy.ChildElements[0] is not A.Position xyPosition ||
                    xy.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefX" or "minX" or "maxX" or "gdRefY" or "minY" or "maxY")) ||
                    xyPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    xyPosition.ChildElements.Count != 0 || xyPosition.X?.Value is not { Length: > 0 } ||
                    xyPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else if (handle is A.AdjustHandlePolar polar)
            {
                if (polar.ChildElements.Count != 1 || polar.ChildElements[0] is not A.Position polarPosition ||
                    polar.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("gdRefR" or "minR" or "maxR" or "gdRefAng" or "minAng" or "maxAng")) ||
                    polarPosition.GetAttributes().Any(attribute =>
                        attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
                    polarPosition.ChildElements.Count != 0 || polarPosition.X?.Value is not { Length: > 0 } ||
                    polarPosition.Y?.Value is not { Length: > 0 })
                    return false;
            }
            else
            {
                return false;
            }
        }
        if (handles[(int)nativeIndex] is not A.AdjustHandlePolar selectedHandle)
            return false;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (shapeHeight is null) return false;
        var position = selectedHandle.GetFirstChild<A.Position>();
        var y = position?.Y?.Value;
        return y is not null && TryLiteralCustomGeometryHandleY(y, shapeHeight.Value, out _);
    }

    private static bool HasSafeNativeShape3dExtrusionHeight(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadExtrusionHeight(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dDepth(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadDepth(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dDepth(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadDepth(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dExtrusionHeight(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadExtrusionHeight(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dContourWidth(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadContourWidth(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dContourRgb(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadContourRgb(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dContourColorScheme(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadContourColorScheme(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dExtrusionRgb(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadExtrusionRgb(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dExtrusionColorScheme(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadExtrusionColorScheme(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneCameraPreset(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraPreset(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneCameraZoom(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraZoom(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneCameraFov(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraFov(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneCameraRotationLatitude(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraRotationLatitude(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneCameraRotationLongitude(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraRotationLongitude(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneCameraRotationRevolution(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraRotationRevolution(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneLightRigPreset(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigPreset(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneLightRigDirection(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigDirection(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneLightRigRotationLatitude(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneLightRigRotationLongitude(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneLightRigRotationRevolution(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropAnchorX(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropAnchorX(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropAnchorY(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropAnchorY(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropAnchorZ(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropAnchorZ(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropNormalDx(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropNormalDx(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropNormalDy(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropNormalDy(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropNormalDz(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropNormalDz(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropUpDx(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropUpDx(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropUpDy(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropUpDy(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dSceneBackdropUpDz(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropUpDz(picture.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativePictureShape3dPresetMaterial(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadPresetMaterial(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dBevelTopWidth(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelTopWidth(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dBevelTopHeight(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelTopHeight(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dBevelTopPreset(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelTopPreset(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dBevelBottomWidth(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelBottomWidth(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dBevelBottomHeight(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelBottomHeight(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativePictureShape3dBevelBottomPreset(P.Picture picture) =>
        picture.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelBottomPreset(picture.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool TryRequestedShape3dExtrusionHeight(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedExtrusionHeight(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedExtrusionHeight(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dExtrusionHeight(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadExtrusionHeight(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadExtrusionHeight(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dContourWidth(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedContourWidth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedContourWidth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dContourWidth(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadContourWidth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadContourWidth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dContourRgb(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedContourRgb(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedContourRgb(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dContourRgb(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadContourRgb(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadContourRgb(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dContourColorScheme(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedContourColorScheme(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedContourColorScheme(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dContourColorScheme(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadContourColorScheme(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadContourColorScheme(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dExtrusionRgb(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedExtrusionRgb(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedExtrusionRgb(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dExtrusionRgb(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadExtrusionRgb(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadExtrusionRgb(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dExtrusionColorScheme(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedExtrusionColorScheme(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedExtrusionColorScheme(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dExtrusionColorScheme(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadExtrusionColorScheme(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadExtrusionColorScheme(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneCameraPreset(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneCameraPreset(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneCameraPreset(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneCameraPreset(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneCameraPreset(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneCameraPreset(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneLightRigPreset(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneLightRigPreset(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneLightRigPreset(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneLightRigPreset(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneLightRigPreset(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneLightRigPreset(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneLightRigDirection(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneLightRigDirection(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneLightRigDirection(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneLightRigDirection(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneLightRigDirection(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneLightRigDirection(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneLightRigRotationLatitude(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneLightRigRotationLatitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneLightRigRotationLatitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneLightRigRotationLatitude(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneLightRigRotationLongitude(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneLightRigRotationLongitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneLightRigRotationLongitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneLightRigRotationLongitude(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneLightRigRotationRevolution(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneLightRigRotationRevolution(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneLightRigRotationRevolution(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneLightRigRotationRevolution(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropAnchorX(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropAnchorX(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropAnchorX(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropAnchorX(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropAnchorX(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropAnchorX(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropAnchorY(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropAnchorY(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropAnchorY(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropAnchorY(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropAnchorY(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropAnchorY(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropAnchorZ(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropAnchorZ(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropAnchorZ(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropAnchorZ(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropAnchorZ(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropAnchorZ(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropNormalDx(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropNormalDx(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropNormalDx(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropNormalDx(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropNormalDx(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropNormalDx(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropNormalDy(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropNormalDy(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropNormalDy(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropNormalDy(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropNormalDy(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropNormalDy(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropNormalDz(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropNormalDz(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropNormalDz(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropNormalDz(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropNormalDz(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropNormalDz(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropUpDx(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropUpDx(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropUpDx(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropUpDx(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropUpDx(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropUpDx(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropUpDy(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropUpDy(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropUpDy(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropUpDy(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropUpDy(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropUpDy(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneBackdropUpDz(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneBackdropUpDz(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneBackdropUpDz(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneBackdropUpDz(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneBackdropUpDz(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneBackdropUpDz(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneCameraZoom(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneCameraZoom(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneCameraZoom(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneCameraZoom(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneCameraZoom(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneCameraZoom(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneCameraFov(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneCameraFov(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneCameraFov(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneCameraFov(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneCameraFov(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneCameraFov(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneCameraRotationLatitude(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneCameraRotationLatitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneCameraRotationLatitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneCameraRotationLatitude(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneCameraRotationLatitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneCameraRotationLatitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneCameraRotationLongitude(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneCameraRotationLongitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneCameraRotationLongitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneCameraRotationLongitude(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneCameraRotationLongitude(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneCameraRotationLongitude(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dSceneCameraRotationRevolution(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedSceneCameraRotationRevolution(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedSceneCameraRotationRevolution(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dSceneCameraRotationRevolution(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadSceneCameraRotationRevolution(
                shape.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadSceneCameraRotationRevolution(
                picture.ShapeProperties?.GetFirstChild<A.Scene3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dPresetMaterial(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedPresetMaterial(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedPresetMaterial(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dPresetMaterial(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadPresetMaterial(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadPresetMaterial(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dBevelTopWidth(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedBevelTopWidth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedBevelTopWidth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dBevelTopWidth(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadBevelTopWidth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadBevelTopWidth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dBevelTopHeight(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedBevelTopHeight(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedBevelTopHeight(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dBevelTopHeight(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadBevelTopHeight(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadBevelTopHeight(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dBevelTopPreset(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedBevelTopPreset(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedBevelTopPreset(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dBevelTopPreset(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadBevelTopPreset(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadBevelTopPreset(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dBevelBottomWidth(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedBevelBottomWidth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedBevelBottomWidth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dBevelBottomWidth(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadBevelBottomWidth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadBevelBottomWidth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dBevelBottomHeight(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedBevelBottomHeight(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedBevelBottomHeight(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dBevelBottomHeight(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadBevelBottomHeight(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadBevelBottomHeight(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dBevelBottomPreset(
        OpenXmlElement element,
        string token,
        out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedBevelBottomPreset(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedBevelBottomPreset(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dBevelBottomPreset(OpenXmlElement element, out string value)
    {
        value = string.Empty;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadBevelBottomPreset(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadBevelBottomPreset(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool TryRequestedShape3dDepth(
        OpenXmlElement element,
        string token,
        out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryRequestedDepth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            P.Picture picture => PptxShape3DCodec.TryRequestedDepth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), token, out value),
            _ => false,
        };
    }

    private static bool TryReadShape3dDepth(OpenXmlElement element, out long value)
    {
        value = 0;
        return element switch
        {
            P.Shape shape => PptxShape3DCodec.TryReadDepth(
                shape.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            P.Picture picture => PptxShape3DCodec.TryReadDepth(
                picture.ShapeProperties?.GetFirstChild<A.Shape3DType>(), out value),
            _ => false,
        };
    }

    private static bool HasSafeNativeShape3dContourWidth(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadContourWidth(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dContourRgb(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadContourRgb(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dContourColorScheme(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadContourColorScheme(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dExtrusionRgb(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadExtrusionRgb(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dExtrusionColorScheme(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadExtrusionColorScheme(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneCameraPreset(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraPreset(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneCameraZoom(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraZoom(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneCameraFov(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraFov(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneCameraRotationLatitude(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraRotationLatitude(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneCameraRotationLongitude(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraRotationLongitude(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneCameraRotationRevolution(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneCameraRotationRevolution(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropAnchorX(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropAnchorX(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropAnchorY(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropAnchorY(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropAnchorZ(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropAnchorZ(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropNormalDx(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropNormalDx(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropNormalDy(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropNormalDy(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropNormalDz(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropNormalDz(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropUpDx(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropUpDx(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropUpDy(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropUpDy(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneBackdropUpDz(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneBackdropUpDz(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneLightRigPreset(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigPreset(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneLightRigDirection(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigDirection(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneLightRigRotationLatitude(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneLightRigRotationLongitude(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dSceneLightRigRotationRevolution(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Scene3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(shape.ShapeProperties.GetFirstChild<A.Scene3DType>(), out _);

    private static bool HasSafeNativeShape3dPresetMaterial(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadPresetMaterial(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dBevelTopWidth(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelTopWidth(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dBevelTopHeight(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelTopHeight(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dBevelTopPreset(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelTopPreset(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dBevelBottomWidth(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelBottomWidth(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dBevelBottomHeight(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelBottomHeight(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeShape3dBevelBottomPreset(P.Shape shape) =>
        shape.ShapeProperties?.Elements<A.Shape3DType>().Count() == 1 &&
        PptxShape3DCodec.TryReadBevelBottomPreset(shape.ShapeProperties.GetFirstChild<A.Shape3DType>(), out _);

    private static bool HasSafeNativeCustomGeometryPathWidth(P.Shape shape, uint nativeIndex) =>
        TryReadNativeCustomGeometryPathWidth(shape, nativeIndex, out _);

    private static bool TryReadNativeCustomGeometryPath(P.Shape shape, uint nativeIndex, out A.Path selected)
    {
        selected = null!;
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var pathList = geometry?.GetFirstChild<A.PathList>();
        if (geometry is null || pathList is null || geometry.HasAttributes || pathList.HasAttributes ||
            geometry.Elements<A.PathList>().Count() != 1 ||
            pathList.ChildElements.Count != pathList.Elements<A.Path>().Count())
            return false;
        var paths = pathList.Elements<A.Path>().ToArray();
        if (nativeIndex >= (uint)paths.Length)
            return false;
        foreach (var path in paths)
        {
            if (path.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 ||
                    attribute.LocalName is not ("w" or "h" or "fill" or "stroke" or "extrusionOk")) ||
                path.ChildElements.Count == 0 ||
                path.ChildElements.Any(command => command is not (A.MoveTo or A.LineTo or A.QuadraticBezierCurveTo or A.CubicBezierCurveTo or A.ArcTo or A.CloseShapePath)))
                return false;
        }
        selected = paths[(int)nativeIndex];
        return true;
    }

    private static bool TryReadNativeCustomGeometryPathWidth(
        P.Shape shape,
        uint nativeIndex,
        out long width)
    {
        width = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            selected.Width?.Value is not { } selectedWidth || selectedWidth is 0 or > int.MaxValue)
            return false;
        width = selectedWidth;
        return true;
    }

    private static bool HasSafeNativeCustomGeometryPathHeight(P.Shape shape, uint nativeIndex) =>
        TryReadNativeCustomGeometryPathHeight(shape, nativeIndex, out _);

    private static bool TryReadNativeCustomGeometryPathHeight(
        P.Shape shape,
        uint nativeIndex,
        out long height)
    {
        height = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            selected.Height?.Value is not { } selectedHeight || selectedHeight is 0 or > int.MaxValue)
            return false;
        height = selectedHeight;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathWidth(P.Shape shape, uint nativeIndex, string token)
    {
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var requested) ||
            requested is <= 0 or > int.MaxValue ||
            requested.ToString(CultureInfo.InvariantCulture) != token ||
            !TryReadNativeCustomGeometryPathWidth(shape, nativeIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool TryRequestedNativeCustomGeometryPathHeight(P.Shape shape, uint nativeIndex, string token)
    {
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var requested) ||
            requested is <= 0 or > int.MaxValue ||
            requested.ToString(CultureInfo.InvariantCulture) != token ||
            !TryReadNativeCustomGeometryPathHeight(shape, nativeIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathArcWidthRadius(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathArcWidthRadius(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathArcWidthRadius(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long widthRadius)
    {
        widthRadius = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.ArcTo arc ||
            arc.ChildElements.Count != 0 ||
            arc.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("wR" or "hR" or "stAng" or "swAng")) ||
            arc.WidthRadius?.Value is not { Length: > 0 } ||
            arc.HeightRadius?.Value is not { Length: > 0 } ||
            arc.StartAngle?.Value is not { Length: > 0 } ||
            arc.SwingAngle?.Value is not { Length: > 0 } ||
            !TryLiteralCustomGeometryPathArcWidthRadius(arc.WidthRadius.Value, out widthRadius))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathArcWidthRadius(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathArcWidthRadius(token, out var requested) ||
            !TryReadNativeCustomGeometryPathArcWidthRadius(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathArcHeightRadius(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathArcHeightRadius(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathArcHeightRadius(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long heightRadius)
    {
        heightRadius = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.ArcTo arc ||
            arc.ChildElements.Count != 0 ||
            arc.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("wR" or "hR" or "stAng" or "swAng")) ||
            arc.WidthRadius?.Value is not { Length: > 0 } ||
            arc.HeightRadius?.Value is not { Length: > 0 } ||
            arc.StartAngle?.Value is not { Length: > 0 } ||
            arc.SwingAngle?.Value is not { Length: > 0 } ||
            !TryLiteralCustomGeometryPathArcHeightRadius(arc.HeightRadius.Value, out heightRadius))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathArcHeightRadius(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathArcHeightRadius(token, out var requested) ||
            !TryReadNativeCustomGeometryPathArcHeightRadius(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathArcStartAngle(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathArcStartAngle(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathArcStartAngle(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long startAngle)
    {
        startAngle = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.ArcTo arc ||
            arc.ChildElements.Count != 0 ||
            arc.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("wR" or "hR" or "stAng" or "swAng")) ||
            arc.WidthRadius?.Value is not { Length: > 0 } ||
            arc.HeightRadius?.Value is not { Length: > 0 } ||
            arc.StartAngle?.Value is not { Length: > 0 } ||
            arc.SwingAngle?.Value is not { Length: > 0 } ||
            !TryLiteralCustomGeometryPathArcStartAngle(arc.StartAngle.Value, out startAngle))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathArcStartAngle(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathArcStartAngle(token, out var requested) ||
            !TryReadNativeCustomGeometryPathArcStartAngle(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathArcSweepAngle(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathArcSweepAngle(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathArcSweepAngle(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long sweepAngle)
    {
        sweepAngle = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.ArcTo arc ||
            arc.ChildElements.Count != 0 ||
            arc.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("wR" or "hR" or "stAng" or "swAng")) ||
            arc.WidthRadius?.Value is not { Length: > 0 } ||
            arc.HeightRadius?.Value is not { Length: > 0 } ||
            arc.StartAngle?.Value is not { Length: > 0 } ||
            arc.SwingAngle?.Value is not { Length: > 0 } ||
            !TryLiteralCustomGeometryPathArcSweepAngle(arc.SwingAngle.Value, out sweepAngle))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathArcSweepAngle(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathArcSweepAngle(token, out var requested) ||
            !TryReadNativeCustomGeometryPathArcSweepAngle(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathLineToX(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathLineToX(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathLineToX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long x)
    {
        x = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.LineTo line ||
            line.ChildElements.Count != 1 ||
            line.ChildElements[0] is not A.Point point ||
            line.HasAttributes ||
            point.ChildElements.Count != 0 ||
            point.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(point.X?.Value, out x))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathLineToX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathLineToX(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathLineToY(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathLineToY(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathLineToY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long y)
    {
        y = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.LineTo line ||
            line.ChildElements.Count != 1 ||
            line.ChildElements[0] is not A.Point point ||
            line.HasAttributes ||
            point.ChildElements.Count != 0 ||
            point.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(point.Y?.Value, out y))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathLineToY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathLineToY(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathMoveToX(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathMoveToX(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathMoveToX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long x)
    {
        x = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.MoveTo moveTo ||
            moveTo.ChildElements.Count != 1 ||
            moveTo.ChildElements[0] is not A.Point point ||
            moveTo.HasAttributes ||
            point.ChildElements.Count != 0 ||
            point.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(point.X?.Value, out x))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathMoveToX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathMoveToX(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathMoveToY(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathMoveToY(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathMoveToY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long y)
    {
        y = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.MoveTo moveTo ||
            moveTo.ChildElements.Count != 1 ||
            moveTo.ChildElements[0] is not A.Point point ||
            moveTo.HasAttributes ||
            point.ChildElements.Count != 0 ||
            point.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(point.Y?.Value, out y))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathMoveToY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathMoveToY(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathQuadraticEndX(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathQuadraticEndX(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathQuadraticEndX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long x)
    {
        x = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.QuadraticBezierCurveTo quadratic ||
            quadratic.HasAttributes ||
            quadratic.ChildElements.Count != 2 ||
            quadratic.ChildElements[0] is not A.Point control ||
            quadratic.ChildElements[1] is not A.Point end ||
            control.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(end.X?.Value, out x))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathQuadraticEndX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathQuadraticEndX(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathQuadraticEndY(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathQuadraticEndY(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathQuadraticEndY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long y)
    {
        y = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.QuadraticBezierCurveTo quadratic ||
            quadratic.HasAttributes ||
            quadratic.ChildElements.Count != 2 ||
            quadratic.ChildElements[0] is not A.Point control ||
            quadratic.ChildElements[1] is not A.Point end ||
            control.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(end.Y?.Value, out y))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathQuadraticEndY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathQuadraticEndY(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathQuadraticControlX(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathQuadraticControlX(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathQuadraticControlX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long x)
    {
        x = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.QuadraticBezierCurveTo quadratic ||
            quadratic.HasAttributes ||
            quadratic.ChildElements.Count != 2 ||
            quadratic.ChildElements[0] is not A.Point control ||
            quadratic.ChildElements[1] is not A.Point end ||
            control.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(control.X?.Value, out x))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathQuadraticControlX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathQuadraticControlX(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathQuadraticControlY(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathQuadraticControlY(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathQuadraticControlY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long y)
    {
        y = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.QuadraticBezierCurveTo quadratic ||
            quadratic.HasAttributes ||
            quadratic.ChildElements.Count != 2 ||
            quadratic.ChildElements[0] is not A.Point control ||
            quadratic.ChildElements[1] is not A.Point end ||
            control.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(control.Y?.Value, out y))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathQuadraticControlY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathQuadraticControlY(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathCubicEndX(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathCubicEndX(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathCubicEndX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long x)
    {
        x = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.CubicBezierCurveTo cubic ||
            cubic.HasAttributes ||
            cubic.ChildElements.Count != 3 ||
            cubic.ChildElements[0] is not A.Point control1 ||
            cubic.ChildElements[1] is not A.Point control2 ||
            cubic.ChildElements[2] is not A.Point end ||
            control1.ChildElements.Count != 0 ||
            control2.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control1.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            control2.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(end.X?.Value, out x))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathCubicEndX(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathCubicEndX(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathCubicEndY(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathCubicEndY(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathCubicEndY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long y)
    {
        y = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.CubicBezierCurveTo cubic ||
            cubic.HasAttributes ||
            cubic.ChildElements.Count != 3 ||
            cubic.ChildElements[0] is not A.Point control1 ||
            cubic.ChildElements[1] is not A.Point control2 ||
            cubic.ChildElements[2] is not A.Point end ||
            control1.ChildElements.Count != 0 ||
            control2.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control1.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            control2.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(end.Y?.Value, out y))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathCubicEndY(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathCubicEndY(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathCubicControl1X(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathCubicControl1X(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathCubicControl1X(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long x)
    {
        x = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.CubicBezierCurveTo cubic ||
            cubic.HasAttributes ||
            cubic.ChildElements.Count != 3 ||
            cubic.ChildElements[0] is not A.Point control1 ||
            cubic.ChildElements[1] is not A.Point control2 ||
            cubic.ChildElements[2] is not A.Point end ||
            control1.ChildElements.Count != 0 ||
            control2.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control1.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            control2.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(control1.X?.Value, out x))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathCubicControl1X(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathCubicControl1X(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathCubicControl1Y(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathCubicControl1Y(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathCubicControl1Y(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long y)
    {
        y = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.CubicBezierCurveTo cubic ||
            cubic.HasAttributes ||
            cubic.ChildElements.Count != 3 ||
            cubic.ChildElements[0] is not A.Point control1 ||
            cubic.ChildElements[1] is not A.Point control2 ||
            cubic.ChildElements[2] is not A.Point end ||
            control1.ChildElements.Count != 0 ||
            control2.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control1.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            control2.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(control1.Y?.Value, out y))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathCubicControl1Y(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathCubicControl1Y(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathCubicControl2X(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathCubicControl2X(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathCubicControl2X(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long x)
    {
        x = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.CubicBezierCurveTo cubic ||
            cubic.HasAttributes ||
            cubic.ChildElements.Count != 3 ||
            cubic.ChildElements[0] is not A.Point control1 ||
            cubic.ChildElements[1] is not A.Point control2 ||
            cubic.ChildElements[2] is not A.Point end ||
            control1.ChildElements.Count != 0 ||
            control2.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control1.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            control2.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(control2.X?.Value, out x))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathCubicControl2X(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathCubicControl2X(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool HasSafeNativeCustomGeometryPathCubicControl2Y(P.Shape shape, uint nativeIndex, uint textIndex) =>
        TryReadNativeCustomGeometryPathCubicControl2Y(shape, nativeIndex, textIndex, out _);

    private static bool TryReadNativeCustomGeometryPathCubicControl2Y(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        out long y)
    {
        y = 0;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) ||
            textIndex >= (uint)selected.ChildElements.Count ||
            selected.ChildElements[(int)textIndex] is not A.CubicBezierCurveTo cubic ||
            cubic.HasAttributes ||
            cubic.ChildElements.Count != 3 ||
            cubic.ChildElements[0] is not A.Point control1 ||
            cubic.ChildElements[1] is not A.Point control2 ||
            cubic.ChildElements[2] is not A.Point end ||
            control1.ChildElements.Count != 0 ||
            control2.ChildElements.Count != 0 ||
            end.ChildElements.Count != 0 ||
            control1.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            control2.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            end.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("x" or "y")) ||
            !TryLiteralCustomGeometryPathCoordinate(control2.Y?.Value, out y))
            return false;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathCubicControl2Y(
        P.Shape shape,
        uint nativeIndex,
        uint textIndex,
        string token)
    {
        if (!TryLiteralCustomGeometryPathCoordinate(token, out var requested) ||
            !TryReadNativeCustomGeometryPathCubicControl2Y(shape, nativeIndex, textIndex, out var current))
            return false;
        return requested != current;
    }

    private static bool TryLiteralCustomGeometryPathArcWidthRadius(string? token, out long radius)
    {
        radius = 0;
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed is <= 0 or > int.MaxValue ||
            parsed.ToString(CultureInfo.InvariantCulture) != token)
            return false;
        radius = parsed;
        return true;
    }

    private static bool TryLiteralCustomGeometryPathArcHeightRadius(string? token, out long radius)
    {
        radius = 0;
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed is <= 0 or > int.MaxValue ||
            parsed.ToString(CultureInfo.InvariantCulture) != token)
            return false;
        radius = parsed;
        return true;
    }

    private static bool TryLiteralCustomGeometryPathArcStartAngle(string? token, out long angle)
    {
        angle = 0;
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < -21_600_000 or > 21_600_000 ||
            parsed.ToString(CultureInfo.InvariantCulture) != token)
            return false;
        angle = parsed;
        return true;
    }

    private static bool TryLiteralCustomGeometryPathArcSweepAngle(string? token, out long angle)
    {
        angle = 0;
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed == 0 || parsed is < -21_600_000 or > 21_600_000 ||
            parsed.ToString(CultureInfo.InvariantCulture) != token)
            return false;
        angle = parsed;
        return true;
    }

    private static bool TryLiteralCustomGeometryPathCoordinate(string? token, out long coordinate)
    {
        coordinate = 0;
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < -int.MaxValue or > int.MaxValue ||
            parsed.ToString(CultureInfo.InvariantCulture) != token)
            return false;
        coordinate = parsed;
        return true;
    }

    private static bool HasSafeNativeCustomGeometryPathFill(P.Shape shape, uint nativeIndex) =>
        TryReadNativeCustomGeometryPathFill(shape, nativeIndex, out _);

    private static bool TryReadNativeCustomGeometryPathFill(P.Shape shape, uint nativeIndex, out bool fill)
    {
        fill = false;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) || selected.Fill?.HasValue != true)
            return false;
        var pathFill = selected.Fill.Value;
        if (pathFill != A.PathFillModeValues.Norm && pathFill != A.PathFillModeValues.None)
            return false;
        fill = pathFill == A.PathFillModeValues.Norm;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathFill(P.Shape shape, uint nativeIndex, string token)
    {
        if (!TryCanonicalBoolean(token, out var requested) ||
            !TryReadNativeCustomGeometryPathFill(shape, nativeIndex, out var current))
            return false;
        return requested != (current ? "1" : "0");
    }

    private static bool HasSafeNativeCustomGeometryPathStroke(P.Shape shape, uint nativeIndex) =>
        TryReadNativeCustomGeometryPathStroke(shape, nativeIndex, out _);

    private static bool TryReadNativeCustomGeometryPathStroke(P.Shape shape, uint nativeIndex, out bool stroke)
    {
        stroke = false;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) || selected.Stroke?.HasValue != true)
            return false;
        stroke = selected.Stroke.Value;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathStroke(P.Shape shape, uint nativeIndex, string token)
    {
        if (!TryCanonicalBoolean(token, out var requested) ||
            !TryReadNativeCustomGeometryPathStroke(shape, nativeIndex, out var current))
            return false;
        return requested != (current ? "1" : "0");
    }

    private static bool HasSafeNativeCustomGeometryPathExtrusionAllowed(P.Shape shape, uint nativeIndex) =>
        TryReadNativeCustomGeometryPathExtrusionAllowed(shape, nativeIndex, out _);

    private static bool TryReadNativeCustomGeometryPathExtrusionAllowed(P.Shape shape, uint nativeIndex, out bool extrusionAllowed)
    {
        extrusionAllowed = false;
        if (!TryReadNativeCustomGeometryPath(shape, nativeIndex, out var selected) || selected.ExtrusionOk?.HasValue != true)
            return false;
        extrusionAllowed = selected.ExtrusionOk.Value;
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryPathExtrusionAllowed(P.Shape shape, uint nativeIndex, string token)
    {
        if (!TryCanonicalBoolean(token, out var requested) ||
            !TryReadNativeCustomGeometryPathExtrusionAllowed(shape, nativeIndex, out var current))
            return false;
        return requested != (current ? "1" : "0");
    }

    private static bool HasSafeNativeCustomGeometryTextRectangleLeft(P.Shape shape) =>
        TryReadNativeCustomGeometryTextRectangle(shape, out _, out _, out _, out _);

    private static bool HasSafeNativeCustomGeometryTextRectangleTop(P.Shape shape) =>
        TryReadNativeCustomGeometryTextRectangle(shape, out _, out _, out _, out _);

    private static bool HasSafeNativeCustomGeometryTextRectangleRight(P.Shape shape) =>
        TryReadNativeCustomGeometryTextRectangle(shape, out _, out _, out _, out _);

    private static bool HasSafeNativeCustomGeometryTextRectangleBottom(P.Shape shape) =>
        TryReadNativeCustomGeometryTextRectangle(shape, out _, out _, out _, out _);

    private static bool TryReadNativeCustomGeometryTextRectangle(
        P.Shape shape,
        out long left,
        out long top,
        out long right,
        out long bottom)
    {
        left = top = right = bottom = 0;
        var geometry = shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var rectangle = geometry?.GetFirstChild<A.Rectangle>();
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        if (geometry is null || rectangle is null || shapeWidth is null || shapeHeight is null ||
            shapeWidth <= 0 || shapeHeight <= 0 || geometry.HasAttributes ||
            geometry.Elements<A.Rectangle>().Count() != 1 || rectangle.ChildElements.Count != 0 ||
            rectangle.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("l" or "t" or "r" or "b")))
            return false;
        var guideList = geometry.GetFirstChild<A.ShapeGuideList>();
        if (guideList is not null &&
            (guideList.HasAttributes || guideList.ChildElements.Count != guideList.Elements<A.ShapeGuide>().Count()))
            return false;
        var guides = guideList?.Elements<A.ShapeGuide>().ToArray() ?? [];
        return TryReadNativeCustomGeometryTextRectangleEdge(rectangle.Left?.Value, "officeKitTextLeft", "w", shapeWidth.Value, guides, out left) &&
            TryReadNativeCustomGeometryTextRectangleEdge(rectangle.Top?.Value, "officeKitTextTop", "h", shapeHeight.Value, guides, out top) &&
            TryReadNativeCustomGeometryTextRectangleEdge(rectangle.Right?.Value, "officeKitTextRight", "w", shapeWidth.Value, guides, out right) &&
            TryReadNativeCustomGeometryTextRectangleEdge(rectangle.Bottom?.Value, "officeKitTextBottom", "h", shapeHeight.Value, guides, out bottom) &&
            left < right && top < bottom;
    }

    private static bool TryReadNativeCustomGeometryTextRectangleEdge(
        string? token,
        string privateGuideName,
        string axis,
        long extent,
        IReadOnlyList<A.ShapeGuide> guides,
        out long coordinate)
    {
        coordinate = 0;
        if (string.IsNullOrEmpty(token) || extent <= 0) return false;
        if (long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var direct) &&
            direct is >= 0 and <= int.MaxValue && direct <= extent &&
            direct.ToString(System.Globalization.CultureInfo.InvariantCulture) == token)
        {
            coordinate = direct;
            return true;
        }
        if (token != privateGuideName) return false;
        var matches = guides.Where(guide => guide.Name?.Value == privateGuideName).ToArray();
        if (matches.Length != 1) return false;
        var guide = matches[0];
        if (guide.ChildElements.Count != 0 ||
            guide.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")))
            return false;
        var formula = guide.Formula?.Value;
        var tokens = (formula ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 4 || tokens[0] != "*/" || tokens[2] != axis ||
            !long.TryParse(tokens[1], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var scaled) ||
            !long.TryParse(tokens[3], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var denominator) ||
            denominator != extent || scaled is < 0 or > int.MaxValue || scaled > extent ||
            scaled.ToString(System.Globalization.CultureInfo.InvariantCulture) != tokens[1] ||
            formula != $"*/ {scaled.ToString(System.Globalization.CultureInfo.InvariantCulture)} {axis} {extent.ToString(System.Globalization.CultureInfo.InvariantCulture)}")
            return false;
        coordinate = scaled;
        return true;
    }

    private static bool TryTextRectanglePrivateGuideFormula(
        string formula,
        string expectedValue,
        out string axis,
        out string extent)
    {
        axis = string.Empty;
        extent = string.Empty;
        var tokens = formula.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 4 || tokens[0] != "*/" || tokens[2] is not ("w" or "h") ||
            !long.TryParse(tokens[1], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var coordinate) ||
            !long.TryParse(tokens[3], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var denominator) ||
            coordinate is < 0 or > int.MaxValue || denominator <= 0 ||
            coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) != expectedValue ||
            formula != $"*/ {coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture)} {tokens[2]} {denominator.ToString(System.Globalization.CultureInfo.InvariantCulture)}")
            return false;
        axis = tokens[2];
        extent = denominator.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryRequestedNativeCustomGeometryTextRectangleLeft(P.Shape shape, string token)
    {
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var requested) ||
            requested is < 0 or > int.MaxValue ||
            requested.ToString(System.Globalization.CultureInfo.InvariantCulture) != token ||
            !TryReadNativeCustomGeometryTextRectangle(shape, out var current, out _, out var right, out _))
            return false;
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        return shapeWidth is not null && requested <= shapeWidth.Value && requested < right && requested != current;
    }

    private static bool TryRequestedNativeCustomGeometryTextRectangleTop(P.Shape shape, string token)
    {
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var requested) ||
            requested is < 0 or > int.MaxValue ||
            requested.ToString(System.Globalization.CultureInfo.InvariantCulture) != token ||
            !TryReadNativeCustomGeometryTextRectangle(shape, out _, out var current, out _, out var bottom))
            return false;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        return shapeHeight is not null && requested <= shapeHeight.Value && requested < bottom && requested != current;
    }

    private static bool TryRequestedNativeCustomGeometryTextRectangleRight(P.Shape shape, string token)
    {
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var requested) ||
            requested is < 0 or > int.MaxValue ||
            requested.ToString(System.Globalization.CultureInfo.InvariantCulture) != token ||
            !TryReadNativeCustomGeometryTextRectangle(shape, out var left, out _, out var current, out _))
            return false;
        var shapeWidth = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
        return shapeWidth is not null && requested <= shapeWidth.Value && requested > left && requested != current;
    }

    private static bool TryRequestedNativeCustomGeometryTextRectangleBottom(P.Shape shape, string token)
    {
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var requested) ||
            requested is < 0 or > int.MaxValue ||
            requested.ToString(System.Globalization.CultureInfo.InvariantCulture) != token ||
            !TryReadNativeCustomGeometryTextRectangle(shape, out _, out var top, out _, out var current))
            return false;
        var shapeHeight = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value;
        return shapeHeight is not null && requested <= shapeHeight.Value && requested > top && requested != current;
    }

    private static bool HasSafeNativePresetGeometryAdjustment(P.Shape shape, uint nativeIndex)
    {
        var geometry = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        if (geometry is null || geometry.Preset?.Value is not { } preset ||
            !PptxCustomGeometryCodec.TryPresetName(preset, out var geometryName) ||
            !PptxPresetGeometryAdjustmentCodec.TryExpectedCount(geometryName, out var expectedCount) ||
            !PptxPresetGeometryAdjustmentCodec.TryReadLiteralSlots(geometry, geometryName, out var values))
            return false;
        return expectedCount > 0 && nativeIndex < (uint)values.Length &&
            !PptxPresetGeometryAdjustmentCodec.IsMissingValue(values[(int)nativeIndex]);
    }

    private static bool TryLiteralCustomAdjustment(string formula, out string value)
    {
        value = string.Empty;
        var tokens = formula.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 2 || tokens[0] != "val" ||
            !long.TryParse(tokens[1], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < int.MinValue or > int.MaxValue)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralConnectionSiteAngle(string token, out string value)
    {
        value = string.Empty;
        if (!int.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < -21_600_000 or > 21_600_000 ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralConnectionSiteX(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralConnectionSiteY(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandleX(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandleY(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandleMinX(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandleMaxX(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandleMinY(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandleMaxY(string token, long extent, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue || parsed > extent ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMinRadius(string token, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMaxRadius(string token, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > int.MaxValue ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMinAngle(string token, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < -21_600_000 or > 21_600_000 ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMaxAngle(string token, out string value)
    {
        value = string.Empty;
        if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < -21_600_000 or > 21_600_000 ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static bool HasSafeNativeShadowGeometry(OpenXmlCompositeElement properties, string kind)
    {
        if (!PptxShadowCodec.TryRead(properties, out var shadow) || shadow is null)
            return false;
        var lists = properties.Elements<A.EffectList>().ToArray();
        if (lists.Length != 1 || lists[0].ChildElements.Count != 1 || lists[0].FirstChild is not A.OuterShadow outer ||
            outer.ChildElements.Count != 1 || outer.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("blurRad" or "dist" or "dir" or "algn" or "rotWithShape")))
            return false;
        var attributeName = kind switch
        {
            "shadowBlurRadiusEmu" or "imageShadowBlurRadiusEmu" => "blurRad",
            "shadowDistanceEmu" or "imageShadowDistanceEmu" => "dist",
            "shadowDirectionDegrees" or "imageShadowDirectionDegrees" => "dir",
            "shadowAlignment" or "imageShadowAlignment" => "algn",
            "shadowRotateWithShape" or "imageShadowRotateWithShape" => "rotWithShape",
            _ => string.Empty,
        };
        var attributes = outer.GetAttributes().Where(attribute => attribute.LocalName == attributeName).ToArray();
        if (attributes.Length != 1 || string.IsNullOrEmpty(attributes[0].Value)) return false;
        if (kind is "shadowBlurRadiusEmu" or "imageShadowBlurRadiusEmu" or "shadowDistanceEmu" or "imageShadowDistanceEmu")
            return long.TryParse(attributes[0].Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var coordinate) &&
                coordinate is >= 0 and <= 2_147_483_647 &&
                coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) == attributes[0].Value;
        if (kind is "shadowDirectionDegrees" or "imageShadowDirectionDegrees")
            return long.TryParse(attributes[0].Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
                direction is >= 0 and < 21_600_000 &&
                direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == attributes[0].Value;
        if (kind is "shadowRotateWithShape" or "imageShadowRotateWithShape")
            return attributes[0].Value is "0" or "1";
        return attributes[0].Value is "tl" or "t" or "tr" or "l" or "ctr" or "r" or "bl" or "b" or "br";
    }

    private static bool HasSafeNativeShadowColor(OpenXmlCompositeElement properties, string kind)
    {
        if (!PptxShadowCodec.TryRead(properties, out var shadow) || shadow is null)
            return false;
        var lists = properties.Elements<A.EffectList>().ToArray();
        if (lists.Length != 1 || lists[0].ChildElements.Count != 1 || lists[0].FirstChild is not A.OuterShadow outer ||
            outer.ChildElements.Count != 1 || outer.FirstChild is not { } color)
            return false;
        if (color.GetAttributes().Count != 1 || color.GetAttributes()[0].NamespaceUri.Length != 0 || color.GetAttributes()[0].LocalName != "val")
            return false;
        var colorMatches = kind is "shadowColorRgb" or "imageShadowColorRgb"
            ? color is A.RgbColorModelHex rgb && rgb.Val?.Value is { Length: 6 } rgbValue && rgbValue.All(Uri.IsHexDigit)
            : color is A.SchemeColor scheme && scheme.Val?.Value is { } schemeValue && PptxColor.TrySchemeToken(schemeValue, out _);
        if (!colorMatches) return false;
        var alphas = color.Elements<A.Alpha>().ToArray();
        return color.ChildElements.Count == alphas.Length && alphas.Length <= 1 &&
            (alphas.Length == 0 || alphas[0].Val?.Value is >= 0 and <= 100_000 &&
                alphas[0].GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "val"));
    }

    private static bool HasSafeNativeShadowOpacity(OpenXmlCompositeElement properties)
    {
        if (!PptxShadowCodec.TryRead(properties, out var shadow) || shadow is null ||
            !shadow.HasOpacityThousandthPercent || shadow.OpacityThousandthPercent > 100_000)
            return false;
        var lists = properties.Elements<A.EffectList>().ToArray();
        if (lists.Length != 1 || lists[0].ChildElements.Count != 1 || lists[0].FirstChild is not A.OuterShadow outer ||
            outer.ChildElements.Count != 1 || outer.FirstChild is not { } color)
            return false;
        var alphas = color.Elements<A.Alpha>().ToArray();
        return color.ChildElements.Count == 1 && alphas.Length == 1 &&
            alphas[0].Val?.Value is >= 0 and <= 100_000 &&
            alphas[0].GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == "val");
    }

    private static bool HasSafeNativeTextGlow(OpenXmlCompositeElement properties) =>
        PptxGlowCodec.TryRead(properties, out var glow) && glow is not null;

    private static bool HasSafeNativeShapeGlow(OpenXmlCompositeElement properties, string kind) =>
        HasSafeNativeGlow(properties, kind);

    private static bool HasSafeNativePictureGlow(P.Picture picture, string kind) =>
        picture.ShapeProperties is { } properties && HasSafeNativeGlow(properties, kind);

    private static bool HasSafeNativePictureShadow(P.Picture picture) =>
        picture.ShapeProperties is { } properties && HasSafeNativeShadowGeometry(properties, "imageShadowRotateWithShape");

    private static bool HasSafeNativePictureShadowGeometry(P.Picture picture, string kind) =>
        picture.ShapeProperties is { } properties && HasSafeNativeShadowGeometry(properties, kind);

    private static bool HasSafeNativePictureShadowColor(P.Picture picture, string kind) =>
        picture.ShapeProperties is { } properties && HasSafeNativeShadowColor(properties, kind);

    private static bool HasSafeNativePictureShadowOpacity(P.Picture picture) =>
        picture.ShapeProperties is { } properties && HasSafeNativeShadowOpacity(properties);

    private static bool HasSafeNativeShapeInnerShadow(OpenXmlCompositeElement properties, string kind) =>
        HasSafeNativeInnerShadow(properties, kind);

    private static bool HasSafeNativePictureInnerShadow(P.Picture picture, string kind) =>
        picture.ShapeProperties is { } properties && HasSafeNativeInnerShadow(properties, kind);

    private static bool HasSafeNativeShapeReflection(OpenXmlCompositeElement properties, string kind) =>
        HasSafeNativeReflection(properties, kind);

    private static bool HasSafeNativePictureReflection(P.Picture picture, string kind) =>
        picture.ShapeProperties is { } properties && HasSafeNativeReflection(properties, kind);

    private static bool HasSafeNativeShapeSoftEdge(OpenXmlCompositeElement properties) =>
        HasSafeNativeSoftEdge(properties);

    private static bool HasSafeNativePictureSoftEdge(P.Picture picture) =>
        picture.ShapeProperties is { } properties && HasSafeNativeSoftEdge(properties);

    private static bool HasSafeNativeSoftEdge(OpenXmlCompositeElement properties)
    {
        if (!PptxSoftEdgeCodec.TryRead(properties, out var softEdge) || softEdge is null)
            return false;
        if (properties.Elements<A.EffectList>().ToArray() is not [{ } effectList] ||
            effectList.ChildElements.LastOrDefault() is not A.SoftEdge nativeSoftEdge ||
            nativeSoftEdge.ChildElements.Count != 0)
            return false;
        return nativeSoftEdge.Radius?.Value is >= 0 and <= 12_700_000;
    }

    private static bool HasSafeNativeGlow(OpenXmlCompositeElement properties, string kind)
    {
        if (!PptxGlowCodec.TryRead(properties, out var glow) || glow is null)
            return false;
        if (properties.Elements<A.EffectList>().ToArray() is not [{ } effectList] ||
            effectList.ChildElements.FirstOrDefault() is not A.Glow nativeGlow ||
            nativeGlow.ChildElements.Count != 1)
            return false;
        return kind switch
        {
            "shapeGlowRadiusEmu" or "imageGlowRadiusEmu" => nativeGlow.Radius?.Value is >= 0 and <= 12_700_000,
            "shapeGlowColorRgb" or "imageGlowColorRgb" => nativeGlow.FirstChild is A.RgbColorModelHex,
            "shapeGlowColorScheme" or "imageGlowColorScheme" => nativeGlow.FirstChild is A.SchemeColor,
            "shapeGlowOpacityThousandthPercent" or "imageGlowOpacityThousandthPercent" =>
                nativeGlow.FirstChild is { } color && color.Elements<A.Alpha>().Count() == 1,
            _ => false,
        };
    }

    private static bool HasSafeNativeInnerShadow(OpenXmlCompositeElement properties, string kind)
    {
        if (!PptxInnerShadowCodec.TryRead(properties, out var shadow) || shadow is null)
            return false;
        if (properties.Elements<A.EffectList>().ToArray() is not [{ } effectList] ||
            effectList.ChildElements.FirstOrDefault() is not A.InnerShadow nativeShadow ||
            nativeShadow.ChildElements.Count != 1)
            return false;
        return kind switch
        {
            "shapeInnerShadowBlurRadiusEmu" or "imageInnerShadowBlurRadiusEmu" => nativeShadow.BlurRadius?.Value is >= 0 and <= int.MaxValue,
            "shapeInnerShadowDistanceEmu" or "imageInnerShadowDistanceEmu" => nativeShadow.Distance?.Value is >= 0 and <= int.MaxValue,
            "shapeInnerShadowDirectionDegrees" or "imageInnerShadowDirectionDegrees" => nativeShadow.Direction?.Value is >= 0 and < 21_600_000,
            "shapeInnerShadowColorRgb" or "imageInnerShadowColorRgb" => nativeShadow.FirstChild is A.RgbColorModelHex,
            "shapeInnerShadowColorScheme" or "imageInnerShadowColorScheme" => nativeShadow.FirstChild is A.SchemeColor,
            "shapeInnerShadowOpacityThousandthPercent" or "imageInnerShadowOpacityThousandthPercent" =>
                nativeShadow.FirstChild is { } color && color.Elements<A.Alpha>().Count() == 1,
            _ => false,
        };
    }

    private static bool HasSafeNativeReflection(OpenXmlCompositeElement properties, string kind)
    {
        if (!PptxReflectionCodec.TryRead(properties, out var reflection) || reflection is null)
            return false;
        if (properties.Elements<A.EffectList>().ToArray() is not [{ } effectList] ||
            effectList.ChildElements.LastOrDefault() is not A.Reflection nativeReflection ||
            nativeReflection.ChildElements.Count != 0)
            return false;
        return kind switch
        {
            "shapeReflectionBlurRadiusEmu" or "imageReflectionBlurRadiusEmu" => nativeReflection.BlurRadius?.Value is >= 0 and <= 12_700_000,
            "shapeReflectionStartOpacityThousandthPercent" or "imageReflectionStartOpacityThousandthPercent" => nativeReflection.StartOpacity?.Value is >= 0 and <= 100_000,
            "shapeReflectionEndOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" => nativeReflection.EndAlpha?.Value is >= 0 and <= 100_000,
            "shapeReflectionDistanceEmu" or "imageReflectionDistanceEmu" => nativeReflection.Distance?.Value is >= 0 and <= 1_270_000_000,
            "shapeReflectionDirectionDegrees" or "imageReflectionDirectionDegrees" => nativeReflection.Direction?.Value is >= 0 and < 21_600_000,
            _ => false,
        };
    }

    private static bool HasSafeNativeTextInnerShadow(OpenXmlCompositeElement properties) =>
        PptxInnerShadowCodec.TryRead(properties, out var shadow) && shadow is not null;

    private static bool HasSafeNativeTextReflection(OpenXmlCompositeElement properties) =>
        PptxReflectionCodec.TryRead(properties, out var reflection) && reflection is not null;

    private static bool HasSafeNativeTextSoftEdge(OpenXmlCompositeElement properties) =>
        PptxSoftEdgeCodec.TryRead(properties, out var softEdge) && softEdge is not null;

    private static bool HasSafeNativePictureOpacity(P.Picture picture)
    {
        var blip = picture.BlipFill?.GetFirstChild<A.Blip>();
        var alpha = blip?.GetFirstChild<A.AlphaModulationFixed>();
        if (alpha is null || alpha.HasChildren || alpha.GetAttributes().Count != 1 ||
            alpha.GetAttributes()[0].LocalName != "amt" || alpha.GetAttributes()[0].NamespaceUri.Length != 0)
            return false;
        return alpha.Amount?.Value is >= 0 and <= 100_000;
    }

    private static bool HasSafeNativePictureMask(P.Picture picture)
    {
        var geometry = picture.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        if (geometry is null || geometry.GetAttributes().Any(attribute => attribute.LocalName != "prst") ||
            geometry.ChildElements.Count != 1 || geometry.FirstChild is not A.AdjustValueList || geometry.FirstChild.ChildElements.Count != 0 ||
            geometry.Preset?.Value is not { } preset)
            return false;
        return PptxCustomGeometryCodec.TryPresetName(preset, out _);
    }

    private static bool HasSafeNativePictureMaskAdjustment(P.Picture picture, uint nativeIndex)
    {
        var geometry = picture.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        if (geometry?.Preset?.Value is not { } preset ||
            !PptxCustomGeometryCodec.TryPresetName(preset, out _) ||
            geometry.GetAttributes().Count != 1 ||
            geometry.GetAttributes()[0].NamespaceUri.Length != 0 ||
            geometry.GetAttributes()[0].LocalName != "prst" ||
            geometry.ChildElements.Count != 1 ||
            geometry.FirstChild is not A.AdjustValueList adjustments ||
            adjustments.HasAttributes ||
            adjustments.ChildElements.Count != adjustments.Elements<A.ShapeGuide>().Count())
            return false;
        var guides = adjustments.Elements<A.ShapeGuide>().ToArray();
        if (nativeIndex >= (uint)guides.Length) return false;
        var target = guides[nativeIndex];
        if (target.ChildElements.Count != 0 || target.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")) ||
            target.Name?.Value is not { Length: > 0 } || target.Formula?.Value is not { } targetFormula)
            return false;
        // Every sibling guide must retain a simple attribute-only shape. Its
        // formula may remain source-owned, but a malformed/child-bearing
        // sibling would make the indexed token splice ambiguous.
        if (guides.Any(guide => guide.ChildElements.Count != 0 ||
                guide.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")) ||
                guide.Name?.Value is not { Length: > 0 } || guide.Formula?.Value is null))
            return false;
        return TryLiteralCustomAdjustment(targetFormula, out _);
    }

    private static bool HasSafeNativeRgbFill(A.SolidFill fill)
    {
        if (fill.ChildElements.Count != 1 || fill.FirstChild is not A.RgbColorModelHex color ||
            color.ChildElements.Count != 0 || color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit))
            return false;
        return true;
    }

    private static bool HasSafeNativeRgbFillOpacity(A.SolidFill fill)
    {
        if (fill.ChildElements.Count != 1 || fill.FirstChild is not A.RgbColorModelHex color ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Alpha alpha ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.GetAttributes().Any(attribute => attribute.LocalName != "val") ||
            alpha.GetAttributes().Any(attribute => attribute.LocalName != "val"))
            return false;
        return alpha.Val?.Value is >= 0 and <= 100_000;
    }

    private static bool HasSafeNativeSchemeFill(A.SolidFill fill)
    {
        if (fill.ChildElements.Count != 1 || fill.FirstChild is not A.SchemeColor color ||
            color.ChildElements.Count != 0 || color.Val?.Value is not { } value ||
            !PptxColor.TrySchemeToken(value, out _))
            return false;
        return color.GetAttributes().All(attribute => attribute.LocalName == "val");
    }

    private static bool HasSafeNativeSchemeFillOpacity(A.SolidFill fill)
    {
        return PptxColor.TryDirectSolidSchemeWithOpacity(fill, out _, out var opacity) && opacity is not null;
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
        if (kind is "tableHeaderRows" or "tableBandedRows" or "tableBandedColumns" or "tableFirstColumnEmphasis" or "tableLastColumnEmphasis" or "tableLastRow")
        {
            if (element is not P.GraphicFrame table ||
                !PptxTableCodec.TryRead(table, out _) ||
                table.Graphic?.GraphicData?.GetFirstChild<A.Table>()?.TableProperties is not { } tableProperties)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct table style flag.", operation.SlidePartPath);
            var value = kind switch
            {
                "tableHeaderRows" => tableProperties.FirstRow?.Value,
                "tableBandedRows" => tableProperties.BandRow?.Value,
                "tableFirstColumnEmphasis" => tableProperties.FirstColumn?.Value,
                "tableLastColumnEmphasis" => tableProperties.LastColumn?.Value,
                "tableLastRow" => tableProperties.LastRow?.Value,
                _ => tableProperties.BandColumn?.Value,
            };
            if (value is not { } flag)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct {kind} flag.", operation.SlidePartPath);
            return flag ? "1" : "0";
        }
        if (kind == "nativeText")
        {
            if ((element is not P.GraphicFrame && element is not P.GroupShape) ||
                !PptxNativeTextLeafCodec.TryResolve(element, operation.TextLeafIndex, out var leaf))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} nativeText target is not a bounded DrawingML text leaf.", operation.SlidePartPath);
            return leaf.Text;
        }
        if (kind == "paragraphAlignment")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraphAlignment target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length ||
                paragraphs[operation.NativeLeafIndex].ParagraphProperties?.Alignment?.Value is not { } alignment ||
                ParagraphAlignmentName(alignment) is not { Length: > 0 } name)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct paragraph alignment.", operation.SlidePartPath);
            return name;
        }
        if (kind is "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" or "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" or "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph spacing index is out of range.", operation.SlidePartPath);
            var paragraphProperties = paragraphs[operation.NativeLeafIndex].ParagraphProperties;
            OpenXmlElement? spacing = kind switch
            {
                "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" => paragraphProperties?.GetFirstChild<A.LineSpacing>(),
                "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" => paragraphProperties?.GetFirstChild<A.SpaceBefore>(),
                "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier" => paragraphProperties?.GetFirstChild<A.SpaceAfter>(),
                _ => null,
            };
            if (spacing is null || spacing.ExtendedAttributes.Any() || spacing.ChildElements.Count != 1)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct paragraph spacing.", operation.SlidePartPath);
            var multiplier = kind.EndsWith("Multiplier", StringComparison.Ordinal);
            if (!multiplier && spacing.GetFirstChild<A.SpacingPoints>()?.Val?.Value is { } points &&
                spacing.FirstChild is A.SpacingPoints pointChild && !pointChild.ExtendedAttributes.Any() &&
                ValidParagraphSpacingNative(points, false, kind is not ("paragraphLineSpacingPoints"), out var pointToken)) return pointToken;
            if (multiplier && spacing.GetFirstChild<A.SpacingPercent>()?.Val?.Value is { } percent &&
                spacing.FirstChild is A.SpacingPercent percentChild && !percentChild.ExtendedAttributes.Any() &&
                ValidParagraphSpacingNative(percent, true, kind is not ("paragraphLineSpacingMultiplier"), out var percentToken)) return percentToken;
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct paragraph spacing of the requested unit.", operation.SlidePartPath);
        }
        if (kind is "paragraphMarginLeftEmu" or "paragraphIndentEmu")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph layout index is out of range.", operation.SlidePartPath);
            var paragraphProperties = paragraphs[operation.NativeLeafIndex].ParagraphProperties;
            var attributeName = kind == "paragraphMarginLeftEmu" ? "marL" : "indent";
            var attributes = paragraphProperties?.GetAttributes()
                .Where(attribute => attribute.LocalName == attributeName)
                .ToArray() ?? [];
            if (attributes.Length != 1)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct paragraph {attributeName}.", operation.SlidePartPath);
            return ParseParagraphLayoutToken(attributes[0].Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph {attributeName} value is missing.", operation.SlidePartPath), kind, operation)
                .ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "paragraphBulletCharacter")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraphBulletCharacter target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph bullet index is out of range.", operation.SlidePartPath);
            var paragraphProperties = paragraphs[operation.NativeLeafIndex].ParagraphProperties;
            var bullets = paragraphProperties?.ChildElements.OfType<A.CharacterBullet>().ToArray() ?? [];
            if (bullets.Length != 1 || !TryReadBulletCharacter(bullets[0], out var character))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct character bullet.", operation.SlidePartPath);
            return character;
        }
        if (kind is "paragraphBulletAutoNumberScheme" or "paragraphBulletAutoNumberStartAt")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph auto-number index is out of range.", operation.SlidePartPath);
            var paragraphProperties = paragraphs[operation.NativeLeafIndex].ParagraphProperties;
            var autoNumbers = paragraphProperties?.ChildElements.OfType<A.AutoNumberedBullet>().ToArray() ?? [];
            if (autoNumbers.Length != 1 || !TryReadAutoNumber(autoNumbers[0], out var scheme, out var startAt))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct auto-number marker.", operation.SlidePartPath);
            return kind == "paragraphBulletAutoNumberScheme"
                ? scheme
                : startAt ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no explicit auto-number start.", operation.SlidePartPath);
        }
        if (kind is "paragraphBulletFontFamily" or "paragraphBulletColorRgb" or "paragraphBulletColorScheme" or "paragraphBulletSizePoints" or "paragraphBulletSizePercent")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph bullet-style index is out of range.", operation.SlidePartPath);
            var paragraphProps = paragraphs[operation.NativeLeafIndex].ParagraphProperties;
            if (paragraphProps is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no direct paragraph properties.", operation.SlidePartPath);
            return ReadParagraphBulletStyleValue(paragraphProps, kind, operation);
        }
        if (kind == "paragraphLevel")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraphLevel target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph level index is out of range.", operation.SlidePartPath);
            var paragraphProperties = paragraphs[operation.NativeLeafIndex].ParagraphProperties;
            var attributes = paragraphProperties?.GetAttributes()
                .Where(attribute => attribute.LocalName == "lvl")
                .ToArray() ?? [];
            if (attributes.Length != 1)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct paragraph level.", operation.SlidePartPath);
            return ParseParagraphLevelToken(attributes[0].Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph level value is missing.", operation.SlidePartPath), operation)
                .ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "verticalAnchor")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} verticalAnchor target is not a shape.", operation.SlidePartPath);
            var anchor = shape.TextBody?.BodyProperties?.Anchor?.Value;
            if (ParagraphVerticalAnchorName(anchor) is not { Length: > 0 } name)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct vertical text anchor.", operation.SlidePartPath);
            return name;
        }
        if (kind is "textBodyInsetLeftEmu" or "textBodyInsetTopEmu" or "textBodyInsetRightEmu" or "textBodyInsetBottomEmu")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var bodyProperties = shape.TextBody?.BodyProperties;
            var value = kind switch
            {
                "textBodyInsetLeftEmu" => bodyProperties?.LeftInset?.Value,
                "textBodyInsetTopEmu" => bodyProperties?.TopInset?.Value,
                "textBodyInsetRightEmu" => bodyProperties?.RightInset?.Value,
                "textBodyInsetBottomEmu" => bodyProperties?.BottomInset?.Value,
                _ => null,
            };
            if (value is not >= 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body inset.", operation.SlidePartPath);
            return value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "textBodyWrap")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyWrap target is not a shape.", operation.SlidePartPath);
            var wrap = shape.TextBody?.BodyProperties?.Wrap?.Value;
            var name = TextBodyWrapName(wrap);
            if (name.Length == 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body wrap mode.", operation.SlidePartPath);
            return name;
        }
        if (kind == "textBodyColumnCount")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyColumnCount target is not a shape.", operation.SlidePartPath);
            var count = shape.TextBody?.BodyProperties?.ColumnCount?.Value;
            if (count is not { } boundedCount || boundedCount is < 1 or > 16)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body column count.", operation.SlidePartPath);
            return boundedCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "textBodyColumnGapEmu")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyColumnGapEmu target is not a shape.", operation.SlidePartPath);
            var gap = shape.TextBody?.BodyProperties?.ColumnSpacing?.Value;
            if (gap is not { } boundedGap || boundedGap < 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body column gap.", operation.SlidePartPath);
            return boundedGap.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "textBodyRotationDegrees")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyRotationDegrees target is not a shape.", operation.SlidePartPath);
            var rotation = shape.TextBody?.BodyProperties?.Rotation?.Value;
            if (rotation is not { } boundedRotation || boundedRotation is < -21_600_000 or > 21_600_000)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body rotation.", operation.SlidePartPath);
            return boundedRotation.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "textBodyVerticalOverflow")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyVerticalOverflow target is not a shape.", operation.SlidePartPath);
            var mode = TextBodyVerticalOverflowName(shape.TextBody?.BodyProperties?.VerticalOverflow?.Value);
            if (mode.Length == 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body vertical overflow.", operation.SlidePartPath);
            return mode;
        }
        if (kind == "textBodyHorizontalOverflow")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyHorizontalOverflow target is not a shape.", operation.SlidePartPath);
            var mode = TextBodyHorizontalOverflowName(shape.TextBody?.BodyProperties?.HorizontalOverflow?.Value);
            if (mode.Length == 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body horizontal overflow.", operation.SlidePartPath);
            return mode;
        }
        if (kind == "textBodyUpright")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyUpright target is not a shape.", operation.SlidePartPath);
            var upright = shape.TextBody?.BodyProperties?.UpRight?.Value;
            if (upright is not { } value)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body upright flag.", operation.SlidePartPath);
            return value ? "1" : "0";
        }
        if (kind == "textBodyAnchorCenter")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyAnchorCenter target is not a shape.", operation.SlidePartPath);
            var anchorCenter = shape.TextBody?.BodyProperties?.AnchorCenter?.Value;
            if (anchorCenter is not { } value)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body anchor-center flag.", operation.SlidePartPath);
            return value ? "1" : "0";
        }
        if (kind == "textBodyForceAntiAlias")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyForceAntiAlias target is not a shape.", operation.SlidePartPath);
            var forceAntiAlias = shape.TextBody?.BodyProperties?.ForceAntiAlias?.Value;
            if (forceAntiAlias is not { } value)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body force anti-alias flag.", operation.SlidePartPath);
            return value ? "1" : "0";
        }
        if (kind == "textBodySpaceFirstLastParagraph")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodySpaceFirstLastParagraph target is not a shape.", operation.SlidePartPath);
            var spaceFirstLastParagraph = shape.TextBody?.BodyProperties?.UseParagraphSpacing?.Value;
            if (spaceFirstLastParagraph is not { } value)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body first-last paragraph spacing flag.", operation.SlidePartPath);
            return value ? "1" : "0";
        }
        if (kind == "textBodyCompatibleLineSpacing")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyCompatibleLineSpacing target is not a shape.", operation.SlidePartPath);
            var compatibleLineSpacing = shape.TextBody?.BodyProperties?.CompatibleLineSpacing?.Value;
            if (compatibleLineSpacing is not { } value)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body compatible line spacing flag.", operation.SlidePartPath);
            return value ? "1" : "0";
        }
        if (kind == "textBodyFromWordArt")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyFromWordArt target is not a shape.", operation.SlidePartPath);
            var fromWordArt = shape.TextBody?.BodyProperties?.FromWordArt?.Value;
            if (fromWordArt is not { } value)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body WordArt marker.", operation.SlidePartPath);
            return value ? "1" : "0";
        }
        if (kind == "textBodyWarpPreset")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyWarpPreset target is not a shape.", operation.SlidePartPath);
            var bodyProperties = shape.TextBody?.BodyProperties;
            if (bodyProperties is null || !PptxBodyPropertiesCodec.TryReadTextWarpPreset(bodyProperties, out var preset))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body warp preset.", operation.SlidePartPath);
            return preset;
        }
        if (kind == "textBodyWarpAdjustment")
        {
            if (element is not P.Shape shape || !HasSafeNativeTextWarpAdjustment(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal text-body warp adjustment.", operation.SlidePartPath);
            var bodyProperties = shape.TextBody!.BodyProperties!;
            var warp = bodyProperties.ChildElements.OfType<A.PresetTextWarp>().Single();
            var guides = warp.GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>().ToArray();
            var formula = guides[(int)operation.NativeLeafIndex].Formula?.Value;
            return formula is not null && PptxBodyPropertiesCodec.TryLiteralTextWarpAdjustment(formula, out var value)
                ? value.ToString(CultureInfo.InvariantCulture)
                : throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal text-body warp adjustment.", operation.SlidePartPath);
        }
        if (kind == "textBodyFlatTextZ")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyFlatTextZ target is not a shape.", operation.SlidePartPath);
            var bodyProperties = shape.TextBody?.BodyProperties;
            if (bodyProperties is null || !PptxBodyPropertiesCodec.TryReadFlatTextZ(bodyProperties, out var coordinate))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body flat-text z coordinate.", operation.SlidePartPath);
            return coordinate.ToString(CultureInfo.InvariantCulture);
        }
        if (kind == "textBodyAutoFit")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyAutoFit target is not a shape.", operation.SlidePartPath);
            var choices = shape.TextBody?.BodyProperties?.ChildElements
                .Where(child => child is A.NoAutoFit or A.NormalAutoFit or A.ShapeAutoFit)
                .ToArray() ?? [];
            if (choices.Length != 1 || !IsBareAutoFitChoice(choices[0]))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body AutoFit mode.", operation.SlidePartPath);
            var mode = TextBodyAutoFitName(choices[0]);
            if (mode.Length == 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body AutoFit mode.", operation.SlidePartPath);
            return mode;
        }
        if (kind is "textBodyNormalAutoFitFontScale" or "textBodyNormalAutoFitLineSpacingReduction")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var bodyProperties = shape.TextBody?.BodyProperties;
            var choices = bodyProperties?.ChildElements
                .Where(child => child is A.NoAutoFit or A.NormalAutoFit or A.ShapeAutoFit)
                .ToArray() ?? [];
            if (choices.Length != 1 || choices[0] is not A.NormalAutoFit normal || normal.ChildElements.Count != 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct normal AutoFit child.", operation.SlidePartPath);
            var attributeName = kind == "textBodyNormalAutoFitFontScale" ? "fontScale" : "lnSpcReduction";
            var attributes = normal.GetAttributes().ToArray();
            if (attributes.Any(attribute => attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("fontScale" or "lnSpcReduction")))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} normal AutoFit child has unsupported attributes.", operation.SlidePartPath);
            var selected = attributes.Where(attribute => attribute.LocalName == attributeName).ToArray();
            if (selected.Length != 1)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct normal AutoFit {attributeName} attribute.", operation.SlidePartPath);
            return ParseTextBodyNormalAutoFitToken(selected[0].Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} normal AutoFit value is missing.", operation.SlidePartPath), kind, operation);
        }
        if (kind == "textBodyColumnDirection")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyColumnDirection target is not a shape.", operation.SlidePartPath);
            var direction = shape.TextBody?.BodyProperties?.RightToLeftColumns?.Value;
            if (direction is not { } value)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body column direction.", operation.SlidePartPath);
            return value ? "1" : "0";
        }
        if (kind == "textBodyVerticalText")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyVerticalText target is not a shape.", operation.SlidePartPath);
            var mode = TextBodyVerticalTextName(shape.TextBody?.BodyProperties?.Vertical?.Value);
            if (mode.Length == 0)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct text-body vertical text mode.", operation.SlidePartPath);
            return mode;
        }
        if (kind == "fontSizePoints")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} fontSizePoints target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run ||
                run.RunProperties?.FontSize?.Value is not { } fontSize || fontSize <= 0 || fontSize > 76_800)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font size.", operation.SlidePartPath);
            return fontSize.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind is "fontFamily" or "fontFamilyEastAsia" or "fontFamilyComplexScript")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font family.", operation.SlidePartPath);
            var runProperties = run.RunProperties;
            OpenXmlElement[] fonts = kind == "fontFamily"
                ? runProperties?.Elements<A.LatinFont>().Cast<OpenXmlElement>().ToArray() ?? []
                : kind == "fontFamilyEastAsia"
                    ? runProperties?.Elements<A.EastAsianFont>().Cast<OpenXmlElement>().ToArray() ?? []
                    : runProperties?.Elements<A.ComplexScriptFont>().Cast<OpenXmlElement>().ToArray() ?? [];
            var typeface = kind == "fontFamily"
                ? (fonts.FirstOrDefault() as A.LatinFont)?.Typeface?.Value
                : kind == "fontFamilyEastAsia"
                    ? (fonts.FirstOrDefault() as A.EastAsianFont)?.Typeface?.Value
                    : (fonts.FirstOrDefault() as A.ComplexScriptFont)?.Typeface?.Value;
            if (fonts.Length != 1 || !ValidFontFamilyToken(typeface ?? string.Empty))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font family.", operation.SlidePartPath);
            return typeface!;
        }
        if (kind is "fontBold" or "fontItalic")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font style.", operation.SlidePartPath);
            var value = kind == "fontBold" ? run.RunProperties?.Bold : run.RunProperties?.Italic;
            if (value is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font style.", operation.SlidePartPath);
            return value.Value ? "1" : "0";
        }
        if (kind is "fontLanguage" or "fontUnderline" or "fontStrike" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run decoration.", operation.SlidePartPath);
            var runProperties = run.RunProperties;
            if (kind == "fontLanguage")
            {
                var language = runProperties?.Language?.Value;
                if (!PptxLanguageTag.IsValid(language))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run language.", operation.SlidePartPath);
                return language!;
            }
            if (kind == "fontUnderline")
            {
                if (!PptxTextDecoration.TryUnderline(runProperties, out var underline))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run underline.", operation.SlidePartPath);
                return underline;
            }
            if (kind == "fontStrike")
            {
                if (!PptxTextDecoration.TryStrike(runProperties, out var strike))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run strike.", operation.SlidePartPath);
                return strike;
            }
            if (kind == "fontBaselinePercent")
            {
                if (!PptxTextDecoration.TryBaseline(runProperties, out var baseline))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run baseline.", operation.SlidePartPath);
                return baseline;
            }
            if (kind == "fontSpacingPoints")
            {
                if (!PptxTextDecoration.TrySpacing(runProperties, out var spacing))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run character spacing.", operation.SlidePartPath);
                return spacing;
            }
            if (kind == "fontCaps")
            {
                if (!PptxTextDecoration.TryCaps(runProperties, out var caps))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run capitalization.", operation.SlidePartPath);
                return caps;
            }
            if (!PptxTextDecoration.TryKerning(runProperties, out var kerning))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run kerning.", operation.SlidePartPath);
            return kerning;
        }
        if (kind is "fontColorRgb" or "fontColorScheme" or "fontHighlightRgb" or "fontHighlightScheme")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font color.", operation.SlidePartPath);
            var isHighlight = kind is "fontHighlightRgb" or "fontHighlightScheme";
            var fill = run.RunProperties?.GetFirstChild<A.SolidFill>();
            string color;
            bool hasColor;
            if (isHighlight)
            {
                var highlightKind = PptxTextDecoration.TryHighlight(run.RunProperties, out var parsedKind, out var parsedValue) ? parsedKind : string.Empty;
                color = parsedValue;
                hasColor = (kind == "fontHighlightRgb" && highlightKind == "rgb") || (kind == "fontHighlightScheme" && highlightKind == "scheme");
            }
            else
            {
                hasColor = kind == "fontColorRgb"
                    ? PptxColor.TryDirectSolidRgb(fill, out color)
                    : PptxColor.TryDirectSolidScheme(fill, out color);
            }
            if (!hasColor)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font color.", operation.SlidePartPath);
            return color;
        }
        if (kind is "textGlowRadiusEmu" or "textGlowColorRgb" or "textGlowColorScheme" or "textGlowOpacityThousandthPercent")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run ||
                run.RunProperties is not { } glowProperties || !HasSafeNativeTextGlow(glowProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text glow.", operation.SlidePartPath);
            if (!PptxGlowCodec.TryRead(glowProperties, out var glow) || glow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text glow.", operation.SlidePartPath);
            return kind switch
            {
                "textGlowRadiusEmu" => glow.RadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textGlowColorRgb" => !string.IsNullOrEmpty(glow.ColorRgb) ? glow.ColorRgb : MissingLeaf(operation),
                "textGlowColorScheme" => !string.IsNullOrEmpty(glow.ColorScheme) ? glow.ColorScheme : MissingLeaf(operation),
                "textGlowOpacityThousandthPercent" => glow.HasOpacityThousandthPercent
                    ? glow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported text glow leaf.", operation.SlidePartPath),
            };
        }
        if (kind is "textDefaultShadowBlurRadiusEmu" or "textDefaultShadowDistanceEmu" or "textDefaultShadowDirectionDegrees" or "textDefaultShadowAlignment" or "textDefaultShadowColorRgb" or "textDefaultShadowColorScheme" or "textDefaultShadowOpacityThousandthPercent" or "textDefaultShadowRotateWithShape")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length ||
                paragraphs[operation.NativeLeafIndex].ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() is not { } shadowProperties ||
                !PptxShadowCodec.TryRead(shadowProperties, out var shadow) || shadow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text outer shadow.", operation.SlidePartPath);
            if (kind == "textDefaultShadowBlurRadiusEmu" && !shadow.HasBlurRadiusEmu ||
                kind == "textDefaultShadowDistanceEmu" && !shadow.HasDistanceEmu ||
                kind == "textDefaultShadowDirectionDegrees" && !shadow.HasDirectionAngle60000 ||
                kind == "textDefaultShadowAlignment" && !shadow.HasAlignment ||
                kind == "textDefaultShadowColorRgb" && string.IsNullOrEmpty(shadow.ColorRgb) ||
                kind == "textDefaultShadowColorScheme" && string.IsNullOrEmpty(shadow.ColorScheme) ||
                kind == "textDefaultShadowOpacityThousandthPercent" && !shadow.HasOpacityThousandthPercent ||
                kind == "textDefaultShadowRotateWithShape" && !shadow.HasRotateWithShape)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text outer shadow {kind} value.", operation.SlidePartPath);
            return kind switch
            {
                "textDefaultShadowBlurRadiusEmu" => shadow.BlurRadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultShadowDistanceEmu" => shadow.DistanceEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultShadowDirectionDegrees" => shadow.DirectionAngle60000.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultShadowAlignment" => shadow.Alignment,
                "textDefaultShadowColorRgb" => shadow.ColorRgb,
                "textDefaultShadowColorScheme" => shadow.ColorScheme,
                "textDefaultShadowOpacityThousandthPercent" => shadow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultShadowRotateWithShape" => shadow.RotateWithShape ? "1" : "0",
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph default-text outer-shadow leaf.", operation.SlidePartPath),
            };
        }
        if (kind is "textDefaultGlowRadiusEmu" or "textDefaultGlowColorRgb" or "textDefaultGlowColorScheme" or "textDefaultGlowOpacityThousandthPercent")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length ||
                paragraphs[operation.NativeLeafIndex].ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() is not { } glowProperties ||
                !HasSafeNativeTextGlow(glowProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text glow.", operation.SlidePartPath);
            if (!PptxGlowCodec.TryRead(glowProperties, out var glow) || glow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text glow.", operation.SlidePartPath);
            return kind switch
            {
                "textDefaultGlowRadiusEmu" => glow.RadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultGlowColorRgb" when !string.IsNullOrEmpty(glow.ColorRgb) => glow.ColorRgb,
                "textDefaultGlowColorScheme" when !string.IsNullOrEmpty(glow.ColorScheme) => glow.ColorScheme,
                "textDefaultGlowOpacityThousandthPercent" when glow.HasOpacityThousandthPercent => glow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture),
                _ => throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded paragraph default-text glow {kind} value.", operation.SlidePartPath),
            };
        }
        if (kind is "shapeGlowRadiusEmu" or "shapeGlowColorRgb" or "shapeGlowColorScheme" or "shapeGlowOpacityThousandthPercent" or
            "imageGlowRadiusEmu" or "imageGlowColorRgb" or "imageGlowColorScheme" or "imageGlowOpacityThousandthPercent")
        {
            var isShape = kind.StartsWith("shapeGlow", StringComparison.Ordinal);
            OpenXmlCompositeElement effectProperties;
            if (isShape && element is P.Shape shape && shape.ShapeProperties is { } shapeProperties)
                effectProperties = shapeProperties;
            else if (!isShape && element is P.Picture picture && picture.ShapeProperties is { } pictureProperties)
                effectProperties = pictureProperties;
            else
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
            if (!HasSafeNativeGlow(effectProperties, kind) || !PptxGlowCodec.TryRead(effectProperties, out var glow) || glow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit glow.", operation.SlidePartPath);
            return kind switch
            {
                "shapeGlowRadiusEmu" or "imageGlowRadiusEmu" => glow.RadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "shapeGlowColorRgb" or "imageGlowColorRgb" => !string.IsNullOrEmpty(glow.ColorRgb) ? glow.ColorRgb : MissingLeaf(operation),
                "shapeGlowColorScheme" or "imageGlowColorScheme" => !string.IsNullOrEmpty(glow.ColorScheme) ? glow.ColorScheme : MissingLeaf(operation),
                "shapeGlowOpacityThousandthPercent" or "imageGlowOpacityThousandthPercent" => glow.HasOpacityThousandthPercent
                    ? glow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported glow leaf.", operation.SlidePartPath),
            };
        }
        if (kind is "shapeInnerShadowBlurRadiusEmu" or "shapeInnerShadowDistanceEmu" or "shapeInnerShadowDirectionDegrees" or "shapeInnerShadowColorRgb" or "shapeInnerShadowColorScheme" or "shapeInnerShadowOpacityThousandthPercent" or
            "imageInnerShadowBlurRadiusEmu" or "imageInnerShadowDistanceEmu" or "imageInnerShadowDirectionDegrees" or "imageInnerShadowColorRgb" or "imageInnerShadowColorScheme" or "imageInnerShadowOpacityThousandthPercent")
        {
            var isShape = kind.StartsWith("shapeInnerShadow", StringComparison.Ordinal);
            OpenXmlCompositeElement effectProperties;
            if (isShape && element is P.Shape shape && shape.ShapeProperties is { } shapeProperties)
                effectProperties = shapeProperties;
            else if (!isShape && element is P.Picture picture && picture.ShapeProperties is { } pictureProperties)
                effectProperties = pictureProperties;
            else
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
            if (!HasSafeNativeInnerShadow(effectProperties, kind) || !PptxInnerShadowCodec.TryRead(effectProperties, out var innerShadow) || innerShadow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit inner shadow.", operation.SlidePartPath);
            return kind switch
            {
                "shapeInnerShadowBlurRadiusEmu" or "imageInnerShadowBlurRadiusEmu" => innerShadow.HasBlurRadiusEmu
                    ? innerShadow.BlurRadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "shapeInnerShadowDistanceEmu" or "imageInnerShadowDistanceEmu" => innerShadow.HasDistanceEmu
                    ? innerShadow.DistanceEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "shapeInnerShadowDirectionDegrees" or "imageInnerShadowDirectionDegrees" => innerShadow.HasDirectionAngle60000
                    ? innerShadow.DirectionAngle60000.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "shapeInnerShadowColorRgb" or "imageInnerShadowColorRgb" => !string.IsNullOrEmpty(innerShadow.ColorRgb) ? innerShadow.ColorRgb : MissingLeaf(operation),
                "shapeInnerShadowColorScheme" or "imageInnerShadowColorScheme" => !string.IsNullOrEmpty(innerShadow.ColorScheme) ? innerShadow.ColorScheme : MissingLeaf(operation),
                "shapeInnerShadowOpacityThousandthPercent" or "imageInnerShadowOpacityThousandthPercent" => innerShadow.HasOpacityThousandthPercent
                    ? innerShadow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported inner-shadow leaf.", operation.SlidePartPath),
            };
        }
        if (kind is "textInnerShadowBlurRadiusEmu" or "textInnerShadowDistanceEmu" or "textInnerShadowDirectionDegrees" or "textInnerShadowColorRgb" or "textInnerShadowColorScheme" or "textInnerShadowOpacityThousandthPercent")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run ||
                run.RunProperties is not { } innerShadowProperties || !HasSafeNativeTextInnerShadow(innerShadowProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text inner shadow.", operation.SlidePartPath);
            if (!PptxInnerShadowCodec.TryRead(innerShadowProperties, out var shadow) || shadow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text inner shadow.", operation.SlidePartPath);
            return kind switch
            {
                "textInnerShadowBlurRadiusEmu" => shadow.HasBlurRadiusEmu
                    ? shadow.BlurRadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "textInnerShadowDistanceEmu" => shadow.HasDistanceEmu
                    ? shadow.DistanceEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "textInnerShadowDirectionDegrees" => shadow.HasDirectionAngle60000
                    ? shadow.DirectionAngle60000.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "textInnerShadowColorRgb" => !string.IsNullOrEmpty(shadow.ColorRgb) ? shadow.ColorRgb : MissingLeaf(operation),
                "textInnerShadowColorScheme" => !string.IsNullOrEmpty(shadow.ColorScheme) ? shadow.ColorScheme : MissingLeaf(operation),
                "textInnerShadowOpacityThousandthPercent" => shadow.HasOpacityThousandthPercent
                    ? shadow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported text inner-shadow leaf.", operation.SlidePartPath),
            };
        }
        if (kind is "textDefaultInnerShadowBlurRadiusEmu" or "textDefaultInnerShadowDistanceEmu" or "textDefaultInnerShadowDirectionDegrees" or "textDefaultInnerShadowColorRgb" or "textDefaultInnerShadowColorScheme" or "textDefaultInnerShadowOpacityThousandthPercent")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textDefaultInnerShadowBlurRadiusEmu target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length ||
                paragraphs[operation.NativeLeafIndex].ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() is not { } innerShadowProperties ||
                !HasSafeNativeTextInnerShadow(innerShadowProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text inner shadow.", operation.SlidePartPath);
            if (!PptxInnerShadowCodec.TryRead(innerShadowProperties, out var shadow) || shadow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text inner shadow.", operation.SlidePartPath);
            return kind switch
            {
                "textDefaultInnerShadowBlurRadiusEmu" when shadow.HasBlurRadiusEmu => shadow.BlurRadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultInnerShadowDistanceEmu" when shadow.HasDistanceEmu => shadow.DistanceEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultInnerShadowDirectionDegrees" when shadow.HasDirectionAngle60000 => shadow.DirectionAngle60000.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultInnerShadowColorRgb" when !string.IsNullOrEmpty(shadow.ColorRgb) => shadow.ColorRgb,
                "textDefaultInnerShadowColorScheme" when !string.IsNullOrEmpty(shadow.ColorScheme) => shadow.ColorScheme,
                "textDefaultInnerShadowOpacityThousandthPercent" when shadow.HasOpacityThousandthPercent => shadow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture),
                _ => throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded paragraph default-text inner shadow {kind} value.", operation.SlidePartPath),
            };
        }
        if (kind is "textDefaultReflectionBlurRadiusEmu" or "textDefaultReflectionDistanceEmu" or "textDefaultReflectionStartOpacityThousandthPercent" or "textDefaultReflectionEndOpacityThousandthPercent" or "textDefaultReflectionDirectionDegrees")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length ||
                paragraphs[operation.NativeLeafIndex].ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() is not { } reflectionProperties ||
                !HasSafeNativeTextReflection(reflectionProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded paragraph default-text reflection.", operation.SlidePartPath);
            if (!PptxReflectionCodec.TryRead(reflectionProperties, out var reflection) || reflection is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded paragraph default-text reflection.", operation.SlidePartPath);
            return kind switch
            {
                "textDefaultReflectionBlurRadiusEmu" when reflection.HasBlurRadiusEmu => reflection.BlurRadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultReflectionDistanceEmu" when reflection.HasDistanceEmu => reflection.DistanceEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultReflectionStartOpacityThousandthPercent" when reflection.HasStartOpacityThousandthPercent => reflection.StartOpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultReflectionEndOpacityThousandthPercent" when reflection.HasEndOpacityThousandthPercent => reflection.EndOpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "textDefaultReflectionDirectionDegrees" when reflection.HasDirectionAngle60000 => reflection.DirectionAngle60000.ToString(System.Globalization.CultureInfo.InvariantCulture),
                _ => throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded paragraph default-text reflection {kind} value.", operation.SlidePartPath),
            };
        }
        if (kind is "shapeReflectionBlurRadiusEmu" or "shapeReflectionStartOpacityThousandthPercent" or "shapeReflectionEndOpacityThousandthPercent" or "shapeReflectionDistanceEmu" or "shapeReflectionDirectionDegrees" or
            "imageReflectionBlurRadiusEmu" or "imageReflectionStartOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" or "imageReflectionDistanceEmu" or "imageReflectionDirectionDegrees")
        {
            var isShape = kind.StartsWith("shapeReflection", StringComparison.Ordinal);
            OpenXmlCompositeElement effectProperties;
            if (isShape && element is P.Shape shape && shape.ShapeProperties is { } shapeProperties)
                effectProperties = shapeProperties;
            else if (!isShape && element is P.Picture picture && picture.ShapeProperties is { } pictureProperties)
                effectProperties = pictureProperties;
            else
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
            if (!HasSafeNativeReflection(effectProperties, kind) || !PptxReflectionCodec.TryRead(effectProperties, out var reflection) || reflection is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit reflection.", operation.SlidePartPath);
            return kind switch
            {
                "shapeReflectionBlurRadiusEmu" or "imageReflectionBlurRadiusEmu" => reflection.HasBlurRadiusEmu
                    ? reflection.BlurRadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "shapeReflectionStartOpacityThousandthPercent" or "imageReflectionStartOpacityThousandthPercent" => reflection.HasStartOpacityThousandthPercent
                    ? reflection.StartOpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "shapeReflectionEndOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" => reflection.HasEndOpacityThousandthPercent
                    ? reflection.EndOpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "shapeReflectionDistanceEmu" or "imageReflectionDistanceEmu" => reflection.HasDistanceEmu
                    ? reflection.DistanceEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "shapeReflectionDirectionDegrees" or "imageReflectionDirectionDegrees" => reflection.HasDirectionAngle60000
                    ? reflection.DirectionAngle60000.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported reflection leaf.", operation.SlidePartPath),
            };
        }
        if (kind is "shapeSoftEdgeRadiusEmu" or "imageSoftEdgeRadiusEmu")
        {
            var isShape = kind == "shapeSoftEdgeRadiusEmu";
            OpenXmlCompositeElement effectProperties;
            if (isShape && element is P.Shape shape && shape.ShapeProperties is { } shapeProperties)
                effectProperties = shapeProperties;
            else if (!isShape && element is P.Picture picture && picture.ShapeProperties is { } pictureProperties)
                effectProperties = pictureProperties;
            else
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
            if (!HasSafeNativeSoftEdge(effectProperties) || !PptxSoftEdgeCodec.TryRead(effectProperties, out var softEdge) || softEdge is null || !softEdge.HasRadiusEmu)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit soft edge.", operation.SlidePartPath);
            return softEdge.RadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind is "textReflectionBlurRadiusEmu" or "textReflectionStartOpacityThousandthPercent" or "textReflectionEndOpacityThousandthPercent" or "textReflectionDistanceEmu" or "textReflectionDirectionDegrees")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run ||
                run.RunProperties is not { } reflectionProperties || !HasSafeNativeTextReflection(reflectionProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text reflection.", operation.SlidePartPath);
            if (!PptxReflectionCodec.TryRead(reflectionProperties, out var reflection) || reflection is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text reflection.", operation.SlidePartPath);
            return kind switch
            {
                "textReflectionBlurRadiusEmu" => reflection.HasBlurRadiusEmu
                    ? reflection.BlurRadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "textReflectionStartOpacityThousandthPercent" => reflection.HasStartOpacityThousandthPercent
                    ? reflection.StartOpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "textReflectionEndOpacityThousandthPercent" => reflection.HasEndOpacityThousandthPercent
                    ? reflection.EndOpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "textReflectionDistanceEmu" => reflection.HasDistanceEmu
                    ? reflection.DistanceEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                "textReflectionDirectionDegrees" => reflection.HasDirectionAngle60000
                    ? reflection.DirectionAngle60000.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : MissingLeaf(operation),
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported text reflection leaf.", operation.SlidePartPath),
            };
        }
        if (kind == "textSoftEdgeRadiusEmu")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textSoftEdgeRadiusEmu target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run ||
                run.RunProperties is not { } softEdgeProperties || !HasSafeNativeTextSoftEdge(softEdgeProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text soft edge.", operation.SlidePartPath);
            if (!PptxSoftEdgeCodec.TryRead(softEdgeProperties, out var softEdge) || softEdge is null || !softEdge.HasRadiusEmu)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit text soft edge.", operation.SlidePartPath);
            return softEdge.RadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "textDefaultSoftEdgeRadiusEmu")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textDefaultSoftEdgeRadiusEmu target is not a shape.", operation.SlidePartPath);
            var paragraphs = shape.Descendants<A.Paragraph>().ToArray();
            if (operation.NativeLeafIndex >= (uint)paragraphs.Length ||
                paragraphs[operation.NativeLeafIndex].ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() is not { } softEdgeProperties ||
                !HasSafeNativeTextSoftEdge(softEdgeProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text soft edge.", operation.SlidePartPath);
            if (!PptxSoftEdgeCodec.TryRead(softEdgeProperties, out var softEdge) || softEdge is null || !softEdge.HasRadiusEmu)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit paragraph default-text soft edge.", operation.SlidePartPath);
            return softEdge.RadiusEmu.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "imageOpacityThousandthPercent")
        {
            if (element is not P.Picture picture || !HasSafeNativePictureOpacity(picture))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image opacity.", operation.SlidePartPath);
            return picture.BlipFill!.GetFirstChild<A.Blip>()!.GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value
                .ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shadowOpacityThousandthPercent")
        {
            if (element is not P.Shape shape || shape.ShapeProperties is null || !HasSafeNativeShadowOpacity(shape.ShapeProperties))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit shadow opacity.", operation.SlidePartPath);
            if (!PptxShadowCodec.TryRead(shape.ShapeProperties, out var shadow) || shadow is null || !shadow.HasOpacityThousandthPercent)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit shadow opacity.", operation.SlidePartPath);
            return shadow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "imageShadowOpacityThousandthPercent")
        {
            if (element is not P.Picture picture || picture.ShapeProperties is null || !HasSafeNativePictureShadowOpacity(picture))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow opacity.", operation.SlidePartPath);
            if (!PptxShadowCodec.TryRead(picture.ShapeProperties, out var shadow) || shadow is null || !shadow.HasOpacityThousandthPercent)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow opacity.", operation.SlidePartPath);
            return shadow.OpacityThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind is "shadowColorRgb" or "shadowColorScheme")
        {
            if (element is not P.Shape shape || shape.ShapeProperties is null || !HasSafeNativeShadowColor(shape.ShapeProperties, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit shadow color.", operation.SlidePartPath);
            if (!PptxShadowCodec.TryRead(shape.ShapeProperties, out var shadow) || shadow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit shadow color.", operation.SlidePartPath);
            var value = kind == "shadowColorRgb" ? shadow.ColorRgb : shadow.ColorScheme;
            return string.IsNullOrEmpty(value) ? MissingLeaf(operation) : value;
        }
        if (kind is "imageShadowColorRgb" or "imageShadowColorScheme")
        {
            if (element is not P.Picture picture || picture.ShapeProperties is null || !HasSafeNativePictureShadowColor(picture, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow color.", operation.SlidePartPath);
            if (!PptxShadowCodec.TryRead(picture.ShapeProperties, out var shadow) || shadow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow color.", operation.SlidePartPath);
            var value = kind == "imageShadowColorRgb" ? shadow.ColorRgb : shadow.ColorScheme;
            return string.IsNullOrEmpty(value) ? MissingLeaf(operation) : value;
        }
        if (kind == "imageMaskPreset")
        {
            if (element is not P.Picture picture || !HasSafeNativePictureMask(picture) ||
                picture.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset?.Value is not { } preset ||
                !PptxCustomGeometryCodec.TryPresetName(preset, out var mask))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image mask preset.", operation.SlidePartPath);
            return mask;
        }
        if (kind == "imageMaskAdjustment")
        {
            if (element is not P.Picture picture || !HasSafeNativePictureMaskAdjustment(picture, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal image mask adjustment.", operation.SlidePartPath);
            var geometry = picture.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!;
            var adjustments = geometry.GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>().ToArray();
            var formula = adjustments[(int)operation.NativeLeafIndex].Formula?.Value;
            return formula is not null && TryLiteralCustomAdjustment(formula, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustment")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustment(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry adjustment.", operation.SlidePartPath);
            var guides = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>().ToArray();
            var formula = guides[(int)operation.NativeLeafIndex].Formula?.Value;
            return formula is not null && TryLiteralCustomAdjustment(formula, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentFormula")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentFormula(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded calculated custom-geometry adjustment formula.", operation.SlidePartPath);
            var guides = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>().ToArray();
            return guides[(int)operation.NativeLeafIndex].Formula?.Value ?? MissingLeaf(operation);
        }
        if (kind == "customGeometryGuide")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryGuide(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry guide.", operation.SlidePartPath);
            var guides = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.ShapeGuideList>()!.Elements<A.ShapeGuide>().ToArray();
            var formula = guides[(int)operation.NativeLeafIndex].Formula?.Value;
            return formula is not null && TryLiteralCustomAdjustment(formula, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryGuideFormula")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryGuideFormula(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded calculated custom-geometry guide formula.", operation.SlidePartPath);
            var guides = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.ShapeGuideList>()!.Elements<A.ShapeGuide>().ToArray();
            return guides[(int)operation.NativeLeafIndex].Formula?.Value ?? MissingLeaf(operation);
        }
        if (kind == "customGeometryConnectionSiteAngle60000")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryConnectionSiteAngle(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry connection-site angle.", operation.SlidePartPath);
            var sites = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.ConnectionSiteList>()!.Elements<A.ConnectionSite>().ToArray();
            var angle = sites[(int)operation.NativeLeafIndex].Angle?.Value;
            return angle is not null && TryLiteralConnectionSiteAngle(angle, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryConnectionSiteXEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryConnectionSiteX(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry connection-site x coordinate.", operation.SlidePartPath);
            var sites = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.ConnectionSiteList>()!.Elements<A.ConnectionSite>().ToArray();
            var position = sites[(int)operation.NativeLeafIndex].GetFirstChild<A.Position>();
            var x = position?.X?.Value;
            var shapeWidth = shape.ShapeProperties.Transform2D?.Extents?.Cx?.Value;
            return x is not null && shapeWidth is not null && TryLiteralConnectionSiteX(x, shapeWidth.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryConnectionSiteYEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryConnectionSiteY(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry connection-site y coordinate.", operation.SlidePartPath);
            var sites = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.ConnectionSiteList>()!.Elements<A.ConnectionSite>().ToArray();
            var position = sites[(int)operation.NativeLeafIndex].GetFirstChild<A.Position>();
            var y = position?.Y?.Value;
            var shapeHeight = shape.ShapeProperties.Transform2D?.Extents?.Cy?.Value;
            return y is not null && shapeHeight is not null && TryLiteralConnectionSiteY(y, shapeHeight.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandleXEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandleX(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry XY adjustment-handle x position.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandleXY)handles[(int)operation.NativeLeafIndex];
            var position = handle.GetFirstChild<A.Position>();
            var x = position?.X?.Value;
            var shapeWidth = shape.ShapeProperties.Transform2D?.Extents?.Cx?.Value;
            return x is not null && shapeWidth is not null && TryLiteralCustomGeometryHandleX(x, shapeWidth.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandleYEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandleY(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry XY adjustment-handle y position.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandleXY)handles[(int)operation.NativeLeafIndex];
            var position = handle.GetFirstChild<A.Position>();
            var y = position?.Y?.Value;
            var shapeHeight = shape.ShapeProperties.Transform2D?.Extents?.Cy?.Value;
            return y is not null && shapeHeight is not null && TryLiteralCustomGeometryHandleY(y, shapeHeight.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandleMinXEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandleMinX(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry XY adjustment-handle minimum-x bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandleXY)handles[(int)operation.NativeLeafIndex];
            var minimum = handle.MinX?.Value;
            var shapeWidth = shape.ShapeProperties.Transform2D?.Extents?.Cx?.Value;
            return minimum is not null && shapeWidth is not null && TryLiteralCustomGeometryHandleMinX(minimum, shapeWidth.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandleMaxXEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandleMaxX(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry XY adjustment-handle maximum-x bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandleXY)handles[(int)operation.NativeLeafIndex];
            var maximum = handle.MaxX?.Value;
            var shapeWidth = shape.ShapeProperties.Transform2D?.Extents?.Cx?.Value;
            return maximum is not null && shapeWidth is not null && TryLiteralCustomGeometryHandleMaxX(maximum, shapeWidth.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandleMinYEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandleMinY(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry XY adjustment-handle minimum-y bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandleXY)handles[(int)operation.NativeLeafIndex];
            var minimum = handle.MinY?.Value;
            var shapeHeight = shape.ShapeProperties.Transform2D?.Extents?.Cy?.Value;
            return minimum is not null && shapeHeight is not null && TryLiteralCustomGeometryHandleMinY(minimum, shapeHeight.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandleMaxYEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandleMaxY(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry XY adjustment-handle maximum-y bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandleXY)handles[(int)operation.NativeLeafIndex];
            var maximum = handle.MaxY?.Value;
            var shapeHeight = shape.ShapeProperties.Transform2D?.Extents?.Cy?.Value;
            return maximum is not null && shapeHeight is not null && TryLiteralCustomGeometryHandleMaxY(maximum, shapeHeight.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandlePolarMinRadiusEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandlePolarMinRadius(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry polar adjustment-handle minimum-radius bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandlePolar)handles[(int)operation.NativeLeafIndex];
            var minimum = handle.MinRadial?.Value;
            return minimum is not null && TryLiteralCustomGeometryHandlePolarMinRadius(minimum, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandlePolarMaxRadiusEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandlePolarMaxRadius(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry polar adjustment-handle maximum-radius bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandlePolar)handles[(int)operation.NativeLeafIndex];
            var maximum = handle.MaxRadial?.Value;
            return maximum is not null && TryLiteralCustomGeometryHandlePolarMaxRadius(maximum, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandlePolarMinAngle60000")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandlePolarMinAngle(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry polar adjustment-handle minimum-angle bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandlePolar)handles[(int)operation.NativeLeafIndex];
            var minimum = handle.MinAngle?.Value;
            return minimum is not null && TryLiteralCustomGeometryHandlePolarMinAngle(minimum, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandlePolarMaxAngle60000")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandlePolarMaxAngle(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry polar adjustment-handle maximum-angle bound.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandlePolar)handles[(int)operation.NativeLeafIndex];
            var maximum = handle.MaxAngle?.Value;
            return maximum is not null && TryLiteralCustomGeometryHandlePolarMaxAngle(maximum, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandlePolarXEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandlePolarX(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry polar adjustment-handle x coordinate.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandlePolar)handles[(int)operation.NativeLeafIndex];
            var shapeWidth = shape.ShapeProperties!.Transform2D?.Extents?.Cx?.Value;
            var position = handle.GetFirstChild<A.Position>();
            var x = position?.X?.Value;
            return x is not null && shapeWidth is not null && TryLiteralCustomGeometryHandleX(x, shapeWidth.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryAdjustmentHandlePolarYEmu")
        {
            if (element is not P.Shape shape || !HasSafeNativeCustomGeometryAdjustmentHandlePolarY(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry polar adjustment-handle y coordinate.", operation.SlidePartPath);
            var handles = shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustHandleList>()!.ChildElements.ToArray();
            var handle = (A.AdjustHandlePolar)handles[(int)operation.NativeLeafIndex];
            var shapeHeight = shape.ShapeProperties!.Transform2D?.Extents?.Cy?.Value;
            var position = handle.GetFirstChild<A.Position>();
            var y = position?.Y?.Value;
            return y is not null && shapeHeight is not null && TryLiteralCustomGeometryHandleY(y, shapeHeight.Value, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind == "customGeometryPathWidth")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathWidth(shape, operation.NativeLeafIndex, out var width))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry path width.", operation.SlidePartPath);
            return width.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathHeight")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathHeight(shape, operation.NativeLeafIndex, out var height))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry path height.", operation.SlidePartPath);
            return height.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathArcWidthRadius")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathArcWidthRadius(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var widthRadius))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry arc width radius.", operation.SlidePartPath);
            return widthRadius.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathArcHeightRadius")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathArcHeightRadius(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var heightRadius))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry arc height radius.", operation.SlidePartPath);
            return heightRadius.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathArcStartAngle60000")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathArcStartAngle(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var startAngle))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry arc start angle.", operation.SlidePartPath);
            return startAngle.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathArcSweepAngle60000")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathArcSweepAngle(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var sweepAngle))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry arc sweep angle.", operation.SlidePartPath);
            return sweepAngle.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dExtrusionHeightEmu")
        {
            if (!TryReadShape3dExtrusionHeight(element, out var extrusionHeight))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D extrusion height.", operation.SlidePartPath);
            return extrusionHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dDepthEmu")
        {
            if (!TryReadShape3dDepth(element, out var depth))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D depth.", operation.SlidePartPath);
            return depth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dContourWidthEmu")
        {
            if (!TryReadShape3dContourWidth(element, out var contourWidth))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D contour width.", operation.SlidePartPath);
            return contourWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dContourRgb")
        {
            if (!TryReadShape3dContourRgb(element, out var contourRgb))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D contour color.", operation.SlidePartPath);
            return contourRgb;
        }
        if (kind == "shape3dContourColorScheme")
        {
            if (!TryReadShape3dContourColorScheme(element, out var contourColorScheme))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D contour theme color.", operation.SlidePartPath);
            return contourColorScheme;
        }
        if (kind == "shape3dExtrusionRgb")
        {
            if (!TryReadShape3dExtrusionRgb(element, out var extrusionRgb))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D extrusion color.", operation.SlidePartPath);
            return extrusionRgb;
        }
        if (kind == "shape3dExtrusionColorScheme")
        {
            if (!TryReadShape3dExtrusionColorScheme(element, out var extrusionColorScheme))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D extrusion theme color.", operation.SlidePartPath);
            return extrusionColorScheme;
        }
        if (kind == "shape3dSceneCameraPreset")
        {
            if (!TryReadShape3dSceneCameraPreset(element, out var sceneCameraPreset))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D scene camera preset.", operation.SlidePartPath);
            return sceneCameraPreset;
        }
        if (kind == "shape3dSceneCameraZoomThousandthPercent")
        {
            if (!TryReadShape3dSceneCameraZoom(element, out var sceneCameraZoom))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D scene camera zoom.", operation.SlidePartPath);
            return sceneCameraZoom.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraFov60000")
        {
            if (!TryReadShape3dSceneCameraFov(element, out var sceneCameraFov))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D scene camera FOV.", operation.SlidePartPath);
            return sceneCameraFov.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraRotationLatitude60000")
        {
            if (!TryReadShape3dSceneCameraRotationLatitude(element, out var sceneCameraRotationLatitude))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D camera rotation latitude.", operation.SlidePartPath);
            return sceneCameraRotationLatitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraRotationLongitude60000")
        {
            if (!TryReadShape3dSceneCameraRotationLongitude(element, out var sceneCameraRotationLongitude))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D camera rotation longitude.", operation.SlidePartPath);
            return sceneCameraRotationLongitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraRotationRevolution60000")
        {
            if (!TryReadShape3dSceneCameraRotationRevolution(element, out var sceneCameraRotationRevolution))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D camera rotation revolution.", operation.SlidePartPath);
            return sceneCameraRotationRevolution.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropAnchorXEmu")
        {
            if (!TryReadShape3dSceneBackdropAnchorX(element, out var sceneBackdropAnchorX))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop anchor X.", operation.SlidePartPath);
            return sceneBackdropAnchorX.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropAnchorYEmu")
        {
            if (!TryReadShape3dSceneBackdropAnchorY(element, out var sceneBackdropAnchorY))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop anchor Y.", operation.SlidePartPath);
            return sceneBackdropAnchorY.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropAnchorZEmu")
        {
            if (!TryReadShape3dSceneBackdropAnchorZ(element, out var sceneBackdropAnchorZ))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop anchor Z.", operation.SlidePartPath);
            return sceneBackdropAnchorZ.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropNormalDxEmu")
        {
            if (!TryReadShape3dSceneBackdropNormalDx(element, out var sceneBackdropNormalDx))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop normal X.", operation.SlidePartPath);
            return sceneBackdropNormalDx.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropNormalDyEmu")
        {
            if (!TryReadShape3dSceneBackdropNormalDy(element, out var sceneBackdropNormalDy))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop normal Y.", operation.SlidePartPath);
            return sceneBackdropNormalDy.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropNormalDzEmu")
        {
            if (!TryReadShape3dSceneBackdropNormalDz(element, out var sceneBackdropNormalDz))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop normal Z.", operation.SlidePartPath);
            return sceneBackdropNormalDz.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropUpDxEmu")
        {
            if (!TryReadShape3dSceneBackdropUpDx(element, out var sceneBackdropUpDx))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop up-vector X.", operation.SlidePartPath);
            return sceneBackdropUpDx.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropUpDyEmu")
        {
            if (!TryReadShape3dSceneBackdropUpDy(element, out var sceneBackdropUpDy))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop up-vector Y.", operation.SlidePartPath);
            return sceneBackdropUpDy.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneBackdropUpDzEmu")
        {
            if (!TryReadShape3dSceneBackdropUpDz(element, out var sceneBackdropUpDz))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D backdrop up-vector Z.", operation.SlidePartPath);
            return sceneBackdropUpDz.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneLightRigPreset")
        {
            if (!TryReadShape3dSceneLightRigPreset(element, out var sceneLightRigPreset))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D scene light-rig preset.", operation.SlidePartPath);
            return sceneLightRigPreset;
        }
        if (kind == "shape3dSceneLightRigDirection")
        {
            if (!TryReadShape3dSceneLightRigDirection(element, out var sceneLightRigDirection))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D scene light-rig direction.", operation.SlidePartPath);
            return sceneLightRigDirection;
        }
        if (kind == "shape3dSceneLightRigRotationLatitude60000")
        {
            if (!TryReadShape3dSceneLightRigRotationLatitude(element, out var sceneLightRigRotationLatitude))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D light-rig rotation latitude.", operation.SlidePartPath);
            return sceneLightRigRotationLatitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneLightRigRotationLongitude60000")
        {
            if (!TryReadShape3dSceneLightRigRotationLongitude(element, out var sceneLightRigRotationLongitude))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D light-rig rotation longitude.", operation.SlidePartPath);
            return sceneLightRigRotationLongitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneLightRigRotationRevolution60000")
        {
            if (!TryReadShape3dSceneLightRigRotationRevolution(element, out var sceneLightRigRotationRevolution))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D light-rig rotation revolution.", operation.SlidePartPath);
            return sceneLightRigRotationRevolution.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dPresetMaterial")
        {
            if (!TryReadShape3dPresetMaterial(element, out var presetMaterial))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D preset material.", operation.SlidePartPath);
            return presetMaterial;
        }
        if (kind == "shape3dBevelTopWidthEmu")
        {
            if (!TryReadShape3dBevelTopWidth(element, out var bevelTopWidth))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D top-bevel width.", operation.SlidePartPath);
            return bevelTopWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelTopHeightEmu")
        {
            if (!TryReadShape3dBevelTopHeight(element, out var bevelTopHeight))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D top-bevel height.", operation.SlidePartPath);
            return bevelTopHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelTopPreset")
        {
            if (!TryReadShape3dBevelTopPreset(element, out var bevelTopPreset))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D top-bevel preset.", operation.SlidePartPath);
            return bevelTopPreset;
        }
        if (kind == "shape3dBevelBottomWidthEmu")
        {
            if (!TryReadShape3dBevelBottomWidth(element, out var bevelBottomWidth))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D bottom-bevel width.", operation.SlidePartPath);
            return bevelBottomWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelBottomHeightEmu")
        {
            if (!TryReadShape3dBevelBottomHeight(element, out var bevelBottomHeight))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D bottom-bevel height.", operation.SlidePartPath);
            return bevelBottomHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelBottomPreset")
        {
            if (!TryReadShape3dBevelBottomPreset(element, out var bevelBottomPreset))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct shape or picture 3-D bottom-bevel preset.", operation.SlidePartPath);
            return bevelBottomPreset;
        }
        if (kind == "customGeometryPathLineToX")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathLineToX(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var x))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry line-to x coordinate.", operation.SlidePartPath);
            return x.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathLineToY")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathLineToY(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var y))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry line-to y coordinate.", operation.SlidePartPath);
            return y.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathMoveToX")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathMoveToX(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var x))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry move-to x coordinate.", operation.SlidePartPath);
            return x.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathMoveToY")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathMoveToY(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var y))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry move-to y coordinate.", operation.SlidePartPath);
            return y.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathQuadraticEndX")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathQuadraticEndX(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var x))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry quadratic end-point x coordinate.", operation.SlidePartPath);
            return x.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathQuadraticEndY")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathQuadraticEndY(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var y))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry quadratic end-point y coordinate.", operation.SlidePartPath);
            return y.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathQuadraticControlX")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathQuadraticControlX(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var x))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry quadratic control-point x coordinate.", operation.SlidePartPath);
            return x.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathQuadraticControlY")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathQuadraticControlY(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var y))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry quadratic control-point y coordinate.", operation.SlidePartPath);
            return y.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathCubicEndX")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathCubicEndX(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var x))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry cubic end-point x coordinate.", operation.SlidePartPath);
            return x.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathCubicEndY")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathCubicEndY(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var y))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry cubic end-point y coordinate.", operation.SlidePartPath);
            return y.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathCubicControl1X")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathCubicControl1X(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var x))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry cubic first-control-point x coordinate.", operation.SlidePartPath);
            return x.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathCubicControl1Y")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathCubicControl1Y(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var y))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry cubic first-control-point y coordinate.", operation.SlidePartPath);
            return y.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathCubicControl2X")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathCubicControl2X(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var x))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry cubic second-control-point x coordinate.", operation.SlidePartPath);
            return x.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathCubicControl2Y")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathCubicControl2Y(shape, operation.NativeLeafIndex, operation.TextLeafIndex, out var y))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry cubic second-control-point y coordinate.", operation.SlidePartPath);
            return y.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryPathFill")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathFill(shape, operation.NativeLeafIndex, out var fill))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no explicit custom-geometry path fill.", operation.SlidePartPath);
            return fill ? "1" : "0";
        }
        if (kind == "customGeometryPathStroke")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathStroke(shape, operation.NativeLeafIndex, out var stroke))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no explicit custom-geometry path stroke.", operation.SlidePartPath);
            return stroke ? "1" : "0";
        }
        if (kind == "customGeometryPathExtrusionAllowed")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryPathExtrusionAllowed(shape, operation.NativeLeafIndex, out var extrusionAllowed))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no explicit custom-geometry path extrusion permission.", operation.SlidePartPath);
            return extrusionAllowed ? "1" : "0";
        }
        if (kind == "customGeometryTextRectangleLeftEmu")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryTextRectangle(shape, out var left, out _, out _, out _))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry text-rectangle left edge.", operation.SlidePartPath);
            return left.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryTextRectangleTopEmu")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryTextRectangle(shape, out _, out var top, out _, out _))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry text-rectangle top edge.", operation.SlidePartPath);
            return top.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryTextRectangleRightEmu")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryTextRectangle(shape, out _, out _, out var right, out _))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry text-rectangle right edge.", operation.SlidePartPath);
            return right.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "customGeometryTextRectangleBottomEmu")
        {
            if (element is not P.Shape shape || !TryReadNativeCustomGeometryTextRectangle(shape, out _, out _, out _, out var bottom))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal custom-geometry text-rectangle bottom edge.", operation.SlidePartPath);
            return bottom.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "presetGeometryAdjustment")
        {
            if (element is not P.Shape shape || !HasSafeNativePresetGeometryAdjustment(shape, operation.NativeLeafIndex))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded literal preset-geometry adjustment.", operation.SlidePartPath);
            var guides = shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!
                .GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>().ToArray();
            var formula = guides[(int)operation.NativeLeafIndex].Formula?.Value;
            return formula is not null && TryLiteralCustomAdjustment(formula, out var value)
                ? value
                : MissingLeaf(operation);
        }
        if (kind is "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment" or "shadowRotateWithShape")
        {
            if (element is not P.Shape shadowShape || shadowShape.ShapeProperties is null ||
                !HasSafeNativeShadowGeometry(shadowShape.ShapeProperties, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit shadow {kind} leaf.", operation.SlidePartPath);
            var effectList = shadowShape.ShapeProperties.Elements<A.EffectList>().Single();
            var outerShadow = (A.OuterShadow)effectList.FirstChild!;
            var attributeName = kind switch
            {
                "shadowBlurRadiusEmu" => "blurRad",
                "shadowDistanceEmu" => "dist",
                "shadowDirectionDegrees" => "dir",
                "shadowAlignment" => "algn",
                "shadowRotateWithShape" => "rotWithShape",
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported shadow geometry.", operation.SlidePartPath),
            };
            var attribute = outerShadow.GetAttributes().Single(item => item.LocalName == attributeName);
            return attribute.Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} shadow geometry attribute has no value.", operation.SlidePartPath);
        }
        if (kind == "imageShadowRotateWithShape")
        {
            if (element is not P.Picture shadowPicture || shadowPicture.ShapeProperties is null ||
                !HasSafeNativePictureShadow(shadowPicture))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow rotation leaf.", operation.SlidePartPath);
            var effectList = shadowPicture.ShapeProperties.Elements<A.EffectList>().Single();
            var outerShadow = (A.OuterShadow)effectList.FirstChild!;
            var attribute = outerShadow.GetAttributes().Single(item => item.LocalName == "rotWithShape");
            return attribute.Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} image shadow rotation attribute has no value.", operation.SlidePartPath);
        }
        if (kind == "imageShadowBlurRadiusEmu")
        {
            if (element is not P.Picture shadowPicture || shadowPicture.ShapeProperties is null ||
                !HasSafeNativePictureShadowGeometry(shadowPicture, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow blur leaf.", operation.SlidePartPath);
            var effectList = shadowPicture.ShapeProperties.Elements<A.EffectList>().Single();
            var outerShadow = (A.OuterShadow)effectList.FirstChild!;
            var attribute = outerShadow.GetAttributes().Single(item => item.LocalName == "blurRad");
            return attribute.Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} image shadow blur attribute has no value.", operation.SlidePartPath);
        }
        if (kind == "imageShadowDistanceEmu")
        {
            if (element is not P.Picture shadowPicture || shadowPicture.ShapeProperties is null ||
                !HasSafeNativePictureShadowGeometry(shadowPicture, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow distance leaf.", operation.SlidePartPath);
            var effectList = shadowPicture.ShapeProperties.Elements<A.EffectList>().Single();
            var outerShadow = (A.OuterShadow)effectList.FirstChild!;
            var attribute = outerShadow.GetAttributes().Single(item => item.LocalName == "dist");
            return attribute.Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} image shadow distance attribute has no value.", operation.SlidePartPath);
        }
        if (kind == "imageShadowDirectionDegrees")
        {
            if (element is not P.Picture shadowPicture || shadowPicture.ShapeProperties is null ||
                !HasSafeNativePictureShadowGeometry(shadowPicture, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow direction leaf.", operation.SlidePartPath);
            var effectList = shadowPicture.ShapeProperties.Elements<A.EffectList>().Single();
            var outerShadow = (A.OuterShadow)effectList.FirstChild!;
            var attribute = outerShadow.GetAttributes().Single(item => item.LocalName == "dir");
            return attribute.Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} image shadow direction attribute has no value.", operation.SlidePartPath);
        }
        if (kind == "imageShadowAlignment")
        {
            if (element is not P.Picture shadowPicture || shadowPicture.ShapeProperties is null ||
                !HasSafeNativePictureShadowGeometry(shadowPicture, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit image shadow alignment leaf.", operation.SlidePartPath);
            var effectList = shadowPicture.ShapeProperties.Elements<A.EffectList>().Single();
            var outerShadow = (A.OuterShadow)effectList.FirstChild!;
            var attribute = outerShadow.GetAttributes().Single(item => item.LocalName == "algn");
            return attribute.Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} image shadow alignment attribute has no value.", operation.SlidePartPath);
        }
        if ((kind is "fillRgb" or "fillScheme" or "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu") && element is P.GroupShape group)
        {
            if (!PptxNativeStyleLeafCodec.TryResolve(group, operation.NativeLeafIndex, out var leaf) || leaf.Kind != kind)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a bounded source-bound group style leaf.", operation.SlidePartPath);
            return leaf.Value;
        }
        if (element is P.GroupShape groupGeometry && (IsGeometryLeaf(kind) || IsGroupChildGeometryLeaf(kind)))
        {
            var groupTransform = groupGeometry.GroupShapeProperties?.GetFirstChild<A.TransformGroup>() ??
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded group transform.", operation.SlidePartPath);
            return kind switch
            {
                "leftEmu" => groupTransform.Offset?.X?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "topEmu" => groupTransform.Offset?.Y?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "widthEmu" => groupTransform.Extents?.Cx?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "heightEmu" => groupTransform.Extents?.Cy?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "childLeftEmu" => groupTransform.ChildOffset?.X?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "childTopEmu" => groupTransform.ChildOffset?.Y?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "childWidthEmu" => groupTransform.ChildExtents?.Cx?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "childHeightEmu" => groupTransform.ChildExtents?.Cy?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "rotationDegrees" => groupTransform.Rotation?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
                "flipHorizontal" => groupTransform.HorizontalFlip?.Value is { } horizontal ? (horizontal ? "1" : "0") : MissingLeaf(operation),
                "flipVertical" => groupTransform.VerticalFlip?.Value is { } vertical ? (vertical ? "1" : "0") : MissingLeaf(operation),
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported group leaf kind {kind}.", operation.SlidePartPath),
            };
        }
        var properties = element switch
        {
            P.Shape shape => shape.ShapeProperties,
            P.ConnectionShape connector when kind is "lineRgb" or "lineScheme" or "lineOpacityThousandthPercent" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineStartArrowWidth" or "lineStartArrowLength" or "lineEndArrowWidth" or "lineEndArrowLength" or "lineWidthEmu" => connector.ShapeProperties,
            P.Picture picture when IsGeometryLeaf(kind) => picture.ShapeProperties,
            _ => null,
        } ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no p:spPr.", operation.SlidePartPath);
        if (element is P.ConnectionShape styleConnector && kind is ("lineRgb" or "lineScheme"))
        {
            var directOutline = properties.GetFirstChild<A.Outline>();
            if (directOutline?.Elements<A.SolidFill>().Any() != true &&
                TryReadNativeConnectorStyleColor(styleConnector, kind, out var styleColor))
                return styleColor;
        }
        var transform = properties.Transform2D;
        return kind switch
        {
            "fillRgb" => RequiredLeafValue(PptxColor.SolidRgb(properties.GetFirstChild<A.SolidFill>()), operation),
            "fillOpacityThousandthPercent" => ReadFillOpacity(properties.GetFirstChild<A.SolidFill>(), operation),
            "fillScheme" => RequiredLeafValue(PptxColor.SolidSchemeWithOpacity(properties.GetFirstChild<A.SolidFill>()), operation),
            "lineRgb" => RequiredLeafValue(PptxColor.SolidRgb(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()), operation),
            "lineScheme" => RequiredLeafValue(NativeSchemeToken(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(), operation), operation),
            "lineOpacityThousandthPercent" => ReadFillOpacity(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(), operation),
            "lineStyle" => RequiredLeafValue(ReadLineStyle(properties.GetFirstChild<A.Outline>(), operation), operation),
            "lineCap" => RequiredLeafValue(ReadLineCap(properties.GetFirstChild<A.Outline>(), operation), operation),
            "lineJoin" => RequiredLeafValue(ReadLineJoin(properties.GetFirstChild<A.Outline>(), operation), operation),
            "lineStartArrow" => RequiredLeafValue(ReadLineArrow(properties.GetFirstChild<A.Outline>(), operation, start: true), operation),
            "lineEndArrow" => RequiredLeafValue(ReadLineArrow(properties.GetFirstChild<A.Outline>(), operation, start: false), operation),
            "lineStartArrowWidth" => RequiredLeafValue(ReadLineArrowSize(properties.GetFirstChild<A.Outline>(), operation, start: true, width: true), operation),
            "lineStartArrowLength" => RequiredLeafValue(ReadLineArrowSize(properties.GetFirstChild<A.Outline>(), operation, start: true, width: false), operation),
            "lineEndArrowWidth" => RequiredLeafValue(ReadLineArrowSize(properties.GetFirstChild<A.Outline>(), operation, start: false, width: true), operation),
            "lineEndArrowLength" => RequiredLeafValue(ReadLineArrowSize(properties.GetFirstChild<A.Outline>(), operation, start: false, width: false), operation),
            "lineWidthEmu" => RequiredLeafValue(properties.GetFirstChild<A.Outline>()?.Width is { } width
                ? width.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : string.Empty, operation),
            "leftEmu" => transform?.Offset?.X?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "topEmu" => transform?.Offset?.Y?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "widthEmu" => transform?.Extents?.Cx?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "heightEmu" => transform?.Extents?.Cy?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "rotationDegrees" => transform?.Rotation?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? MissingLeaf(operation),
            "flipHorizontal" => transform?.HorizontalFlip?.Value is { } horizontal ? (horizontal ? "1" : "0") : MissingLeaf(operation),
            "flipVertical" => transform?.VerticalFlip?.Value is { } vertical ? (vertical ? "1" : "0") : MissingLeaf(operation),
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported leaf kind {kind}.", operation.SlidePartPath),
        };
    }

    private static string RequiredLeafValue(string value, PresentationEditOperation operation) =>
        string.IsNullOrEmpty(value) ? MissingLeaf(operation) : value;

    private static string ReadFillOpacity(A.SolidFill? fill, PresentationEditOperation operation)
    {
        if (PptxColor.TryDirectSolidRgbWithOpacity(fill, out _, out var rgbOpacity) && rgbOpacity is not null)
            return rgbOpacity.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (PptxColor.TryDirectSolidSchemeWithOpacity(fill, out _, out var schemeOpacity) && schemeOpacity is not null)
            return schemeOpacity.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return MissingLeaf(operation);
    }

    private static string NativeSchemeToken(A.SolidFill? fill, PresentationEditOperation operation)
    {
        var scheme = fill?.GetFirstChild<A.SchemeColor>()?.Val?.Value;
        return scheme is { } value && PptxColor.TrySchemeToken(value, out var token)
            ? token
            : MissingLeaf(operation);
    }

    private static string ReadLineStyle(A.Outline? outline, PresentationEditOperation operation) =>
        outline?.GetFirstChild<A.PresetDash>() is { } dash && PptxLineStyleCodec.TryReadPresetDash(dash, out var style)
            ? style
            : MissingLeaf(operation);

    private static string ReadLineCap(A.Outline? outline, PresentationEditOperation operation) =>
        PptxLineStyleCodec.TryReadCap(outline, out var cap)
            ? cap
            : MissingLeaf(operation);

    private static string ReadLineJoin(A.Outline? outline, PresentationEditOperation operation) =>
        PptxLineStyleCodec.TryReadJoinLeaf(outline, out var join)
            ? join
            : MissingLeaf(operation);

    private static string ReadLineArrow(A.Outline? outline, PresentationEditOperation operation, bool start) =>
        PptxLineStyleCodec.TryReadArrowType(start ? outline?.GetFirstChild<A.HeadEnd>() : outline?.GetFirstChild<A.TailEnd>(), out var arrow)
            ? arrow
            : MissingLeaf(operation);

    private static string ReadLineArrowSize(A.Outline? outline, PresentationEditOperation operation, bool start, bool width) =>
        PptxLineStyleCodec.TryReadArrowSize(start ? outline?.GetFirstChild<A.HeadEnd>() : outline?.GetFirstChild<A.TailEnd>(), width, out var size)
            ? size
            : MissingLeaf(operation);

    private static string ParagraphAlignmentName(A.TextAlignmentTypeValues value) =>
        value == A.TextAlignmentTypeValues.Left ? "left" :
        value == A.TextAlignmentTypeValues.Center ? "center" :
        value == A.TextAlignmentTypeValues.Right ? "right" :
        value == A.TextAlignmentTypeValues.Justified ? "justify" :
        value == A.TextAlignmentTypeValues.Distributed ? "distributed" : string.Empty;

    private static string ParagraphAlignmentName(string value) =>
        value switch
        {
            "l" => "left",
            "ctr" => "center",
            "r" => "right",
            "just" => "justify",
            "dist" => "distributed",
            _ => string.Empty,
        };

    private static string ParagraphAlignmentToken(string value) =>
        value switch
        {
            "left" => "l",
            "center" => "ctr",
            "right" => "r",
            "justify" => "just",
            _ => throw new CodecException("invalid_presentation_edit_target", $"Unsupported Presentation paragraph alignment {value}."),
        };

    private static string ParagraphVerticalAnchorName(A.TextAnchoringTypeValues? value)
    {
        if (value is null) return string.Empty;
        if (value.Value == A.TextAnchoringTypeValues.Top) return "top";
        if (value.Value == A.TextAnchoringTypeValues.Center) return "center";
        if (value.Value == A.TextAnchoringTypeValues.Bottom) return "bottom";
        return string.Empty;
    }

    private static string ParagraphVerticalAnchorName(string value) => value switch
    {
        "t" => "top",
        "ctr" => "center",
        "b" => "bottom",
        _ => string.Empty,
    };

    private static string ParagraphVerticalAnchorToken(string value) => value switch
    {
        "top" => "t",
        "center" => "ctr",
        "bottom" => "b",
        _ => throw new CodecException("invalid_presentation_edit_target", $"Unsupported Presentation vertical anchor {value}."),
    };

    private static string MissingLeaf(PresentationEditOperation operation) =>
        throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no {LeafKind(operation)} leaf.", operation.SlidePartPath);

    private static PptxXmlPatch CompileScalarXmlPatch(string xml, XmlRange elementRange, PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var owner = elementRange.LocalName;
        if (LeafKind(operation) is "tableHeaderRows" or "tableBandedRows" or "tableBandedColumns" or "tableFirstColumnEmphasis" or "tableLastColumnEmphasis" or "tableLastRow")
            return CompileTableBooleanXmlPatch(xml, elementRange, proof);
        if (LeafKind(operation) is "shapeGlowRadiusEmu" or "shapeGlowColorRgb" or "shapeGlowColorScheme" or "shapeGlowOpacityThousandthPercent" or
            "imageGlowRadiusEmu" or "imageGlowColorRgb" or "imageGlowColorScheme" or "imageGlowOpacityThousandthPercent")
            return CompileGlowXmlPatch(xml, elementRange, proof);
        if (LeafKind(operation) is "shapeInnerShadowBlurRadiusEmu" or "shapeInnerShadowDistanceEmu" or "shapeInnerShadowDirectionDegrees" or "shapeInnerShadowColorRgb" or "shapeInnerShadowColorScheme" or "shapeInnerShadowOpacityThousandthPercent" or
            "imageInnerShadowBlurRadiusEmu" or "imageInnerShadowDistanceEmu" or "imageInnerShadowDirectionDegrees" or "imageInnerShadowColorRgb" or "imageInnerShadowColorScheme" or "imageInnerShadowOpacityThousandthPercent")
            return CompileInnerShadowXmlPatch(xml, elementRange, proof);
        if (LeafKind(operation) is "shapeReflectionBlurRadiusEmu" or "shapeReflectionStartOpacityThousandthPercent" or "shapeReflectionEndOpacityThousandthPercent" or "shapeReflectionDistanceEmu" or "shapeReflectionDirectionDegrees" or
            "imageReflectionBlurRadiusEmu" or "imageReflectionStartOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" or "imageReflectionDistanceEmu" or "imageReflectionDirectionDegrees")
            return CompileReflectionXmlPatch(xml, elementRange, proof);
        if (LeafKind(operation) is "shapeSoftEdgeRadiusEmu" or "imageSoftEdgeRadiusEmu")
            return CompileSoftEdgeXmlPatch(xml, elementRange, proof);
        if (owner == "grpSp" && (LeafKind(operation) is "fillRgb" or "fillScheme" or "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu"))
            return CompileNativeStyleXmlPatch(xml, elementRange, proof);
        var propertiesName = owner == "grpSp" ? "grpSpPr" : "spPr";
        var properties = DirectChildRange(xml, elementRange, owner, propertiesName, operation);
        XmlRange leaf;
        string attribute;
        switch (LeafKind(operation))
        {
            case "fillRgb":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} picture targets do not expose fillRgb.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, "spPr", "solidFill", operation), "solidFill", "srgbClr", operation);
                attribute = "val";
                break;
            case "fillOpacityThousandthPercent":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose fillOpacityThousandthPercent.", operation.SlidePartPath);
                var solidFill = DirectChildRange(xml, properties, "spPr", "solidFill", operation);
                var solidColors = DirectChildRanges(xml, solidFill);
                if (solidColors.Count != 1 || solidColors[0].LocalName is not ("srgbClr" or "schemeClr"))
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct RGB or theme color under solidFill.", operation.SlidePartPath);
                var solidColor = solidColors[0];
                leaf = DirectChildRange(xml, solidColor, solidColor.LocalName, "alpha", operation);
                attribute = "val";
                break;
            case "shape3dExtrusionHeightEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D extrusion height.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                attribute = "extrusionH";
                break;
            case "shape3dDepthEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D depth.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                attribute = "z";
                break;
            case "shape3dContourWidthEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D contour width.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                attribute = "contourW";
                break;
            case "shape3dContourRgb":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D contour color.", operation.SlidePartPath);
                var contourColorShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                var contourColorOwner = DirectChildRange(xml, contourColorShape3d, "sp3d", "contourClr", operation);
                leaf = DirectChildRange(xml, contourColorOwner, "contourClr", "srgbClr", operation);
                attribute = "val";
                break;
            case "shape3dContourColorScheme":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D contour theme color.", operation.SlidePartPath);
                var contourSchemeShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                var contourSchemeOwner = DirectChildRange(xml, contourSchemeShape3d, "sp3d", "contourClr", operation);
                leaf = DirectChildRange(xml, contourSchemeOwner, "contourClr", "schemeClr", operation);
                attribute = "val";
                break;
            case "shape3dExtrusionRgb":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D extrusion color.", operation.SlidePartPath);
                var extrusionColorShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                var extrusionColorOwner = DirectChildRange(xml, extrusionColorShape3d, "sp3d", "extrusionClr", operation);
                leaf = DirectChildRange(xml, extrusionColorOwner, "extrusionClr", "srgbClr", operation);
                attribute = "val";
                break;
            case "shape3dExtrusionColorScheme":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D extrusion theme color.", operation.SlidePartPath);
                var extrusionSchemeShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                var extrusionSchemeOwner = DirectChildRange(xml, extrusionSchemeShape3d, "sp3d", "extrusionClr", operation);
                leaf = DirectChildRange(xml, extrusionSchemeOwner, "extrusionClr", "schemeClr", operation);
                attribute = "val";
                break;
            case "shape3dSceneCameraPreset":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D scene camera preset.", operation.SlidePartPath);
                var sceneCameraOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                leaf = DirectChildRange(xml, sceneCameraOwner, "scene3d", "camera", operation);
                attribute = "prst";
                break;
            case "shape3dSceneCameraZoomThousandthPercent":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D scene camera zoom.", operation.SlidePartPath);
                var sceneCameraZoomOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                leaf = DirectChildRange(xml, sceneCameraZoomOwner, "scene3d", "camera", operation);
                attribute = "zoom";
                break;
            case "shape3dSceneCameraFov60000":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D scene camera FOV.", operation.SlidePartPath);
                var sceneCameraFovOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                leaf = DirectChildRange(xml, sceneCameraFovOwner, "scene3d", "camera", operation);
                attribute = "fov";
                break;
            case "shape3dSceneCameraRotationLatitude60000":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D camera rotation latitude.", operation.SlidePartPath);
                var sceneCameraRotationLatitudeOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneCameraRotationLatitude = DirectChildRange(xml, sceneCameraRotationLatitudeOwner, "scene3d", "camera", operation);
                leaf = DirectChildRange(xml, sceneCameraRotationLatitude, "camera", "rot", operation);
                attribute = "lat";
                break;
            case "shape3dSceneCameraRotationLongitude60000":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D camera rotation longitude.", operation.SlidePartPath);
                var sceneCameraRotationLongitudeOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneCameraRotationLongitude = DirectChildRange(xml, sceneCameraRotationLongitudeOwner, "scene3d", "camera", operation);
                leaf = DirectChildRange(xml, sceneCameraRotationLongitude, "camera", "rot", operation);
                attribute = "lon";
                break;
            case "shape3dSceneCameraRotationRevolution60000":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D camera rotation revolution.", operation.SlidePartPath);
                var sceneCameraRotationRevolutionOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneCameraRotationRevolution = DirectChildRange(xml, sceneCameraRotationRevolutionOwner, "scene3d", "camera", operation);
                leaf = DirectChildRange(xml, sceneCameraRotationRevolution, "camera", "rot", operation);
                attribute = "rev";
                break;
            case "shape3dSceneBackdropAnchorXEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop anchor X.", operation.SlidePartPath);
                var sceneBackdropAnchorOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdrop = DirectChildRange(xml, sceneBackdropAnchorOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdrop, "backdrop", "anchor", operation);
                attribute = "x";
                break;
            case "shape3dSceneBackdropAnchorYEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop anchor Y.", operation.SlidePartPath);
                var sceneBackdropAnchorYOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropY = DirectChildRange(xml, sceneBackdropAnchorYOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropY, "backdrop", "anchor", operation);
                attribute = "y";
                break;
            case "shape3dSceneBackdropAnchorZEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop anchor Z.", operation.SlidePartPath);
                var sceneBackdropAnchorZOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropZ = DirectChildRange(xml, sceneBackdropAnchorZOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropZ, "backdrop", "anchor", operation);
                attribute = "z";
                break;
            case "shape3dSceneBackdropNormalDxEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop normal X.", operation.SlidePartPath);
                var sceneBackdropNormalDxOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropNormal = DirectChildRange(xml, sceneBackdropNormalDxOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropNormal, "backdrop", "norm", operation);
                attribute = "dx";
                break;
            case "shape3dSceneBackdropNormalDyEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop normal Y.", operation.SlidePartPath);
                var sceneBackdropNormalDyOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropNormalDy = DirectChildRange(xml, sceneBackdropNormalDyOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropNormalDy, "backdrop", "norm", operation);
                attribute = "dy";
                break;
            case "shape3dSceneBackdropNormalDzEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop normal Z.", operation.SlidePartPath);
                var sceneBackdropNormalDzOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropNormalDz = DirectChildRange(xml, sceneBackdropNormalDzOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropNormalDz, "backdrop", "norm", operation);
                attribute = "dz";
                break;
            case "shape3dSceneBackdropUpDxEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop up-vector X.", operation.SlidePartPath);
                var sceneBackdropUpDxOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropUp = DirectChildRange(xml, sceneBackdropUpDxOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropUp, "backdrop", "up", operation);
                attribute = "dx";
                break;
            case "shape3dSceneBackdropUpDyEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop up-vector Y.", operation.SlidePartPath);
                var sceneBackdropUpDyOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropUpDy = DirectChildRange(xml, sceneBackdropUpDyOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropUpDy, "backdrop", "up", operation);
                attribute = "dy";
                break;
            case "shape3dSceneBackdropUpDzEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D backdrop up-vector Z.", operation.SlidePartPath);
                var sceneBackdropUpDzOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneBackdropUpDz = DirectChildRange(xml, sceneBackdropUpDzOwner, "scene3d", "backdrop", operation);
                leaf = DirectChildRange(xml, sceneBackdropUpDz, "backdrop", "up", operation);
                attribute = "dz";
                break;
            case "shape3dSceneLightRigPreset":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D scene light-rig preset.", operation.SlidePartPath);
                var sceneLightRigOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                leaf = DirectChildRange(xml, sceneLightRigOwner, "scene3d", "lightRig", operation);
                attribute = "rig";
                break;
            case "shape3dSceneLightRigDirection":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D scene light-rig direction.", operation.SlidePartPath);
                var sceneLightRigDirectionOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                leaf = DirectChildRange(xml, sceneLightRigDirectionOwner, "scene3d", "lightRig", operation);
                attribute = "dir";
                break;
            case "shape3dSceneLightRigRotationLongitude60000":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D light-rig rotation longitude.", operation.SlidePartPath);
                var sceneLightRigRotationLongitudeOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneLightRigRotationLongitude = DirectChildRange(xml, sceneLightRigRotationLongitudeOwner, "scene3d", "lightRig", operation);
                leaf = DirectChildRange(xml, sceneLightRigRotationLongitude, "lightRig", "rot", operation);
                attribute = "lon";
                break;
            case "shape3dSceneLightRigRotationRevolution60000":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D light-rig rotation revolution.", operation.SlidePartPath);
                var sceneLightRigRotationRevolutionOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneLightRigRotationRevolution = DirectChildRange(xml, sceneLightRigRotationRevolutionOwner, "scene3d", "lightRig", operation);
                leaf = DirectChildRange(xml, sceneLightRigRotationRevolution, "lightRig", "rot", operation);
                attribute = "rev";
                break;
            case "shape3dSceneLightRigRotationLatitude60000":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D light-rig rotation latitude.", operation.SlidePartPath);
                var sceneLightRigRotationOwner = DirectChildRange(xml, properties, "spPr", "scene3d", operation);
                var sceneLightRigRotation = DirectChildRange(xml, sceneLightRigRotationOwner, "scene3d", "lightRig", operation);
                leaf = DirectChildRange(xml, sceneLightRigRotation, "lightRig", "rot", operation);
                attribute = "lat";
                break;
            case "shape3dPresetMaterial":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D preset material.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                attribute = "prstMaterial";
                break;
            case "shape3dBevelTopWidthEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D top-bevel width.", operation.SlidePartPath);
                var bevelShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                leaf = DirectChildRange(xml, bevelShape3d, "sp3d", "bevelT", operation);
                attribute = "w";
                break;
            case "shape3dBevelTopHeightEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D top-bevel height.", operation.SlidePartPath);
                var bevelHeightShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                leaf = DirectChildRange(xml, bevelHeightShape3d, "sp3d", "bevelT", operation);
                attribute = "h";
                break;
            case "shape3dBevelTopPreset":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D top-bevel preset.", operation.SlidePartPath);
                var bevelPresetShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                leaf = DirectChildRange(xml, bevelPresetShape3d, "sp3d", "bevelT", operation);
                attribute = "prst";
                break;
            case "shape3dBevelBottomWidthEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D bottom-bevel width.", operation.SlidePartPath);
                var bevelBottomShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                leaf = DirectChildRange(xml, bevelBottomShape3d, "sp3d", "bevelB", operation);
                attribute = "w";
                break;
            case "shape3dBevelBottomHeightEmu":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D bottom-bevel height.", operation.SlidePartPath);
                var bevelBottomHeightShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                leaf = DirectChildRange(xml, bevelBottomHeightShape3d, "sp3d", "bevelB", operation);
                attribute = "h";
                break;
            case "shape3dBevelBottomPreset":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shape or picture 3-D bottom-bevel preset.", operation.SlidePartPath);
                var bevelBottomPresetShape3d = DirectChildRange(xml, properties, "spPr", "sp3d", operation);
                leaf = DirectChildRange(xml, bevelBottomPresetShape3d, "sp3d", "bevelB", operation);
                attribute = "prst";
                break;
            case "shadowBlurRadiusEmu":
            case "shadowDistanceEmu":
            case "shadowDirectionDegrees":
            case "shadowAlignment":
            case "shadowRotateWithShape":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var shadowEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                leaf = DirectChildRange(xml, shadowEffectList, "effectLst", "outerShdw", operation);
                attribute = LeafKind(operation) switch
                {
                    "shadowBlurRadiusEmu" => "blurRad",
                    "shadowDistanceEmu" => "dist",
                    "shadowDirectionDegrees" => "dir",
                    "shadowAlignment" => "algn",
                    "shadowRotateWithShape" => "rotWithShape",
                    _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported shadow geometry.", operation.SlidePartPath),
                };
                break;
            case "imageShadowRotateWithShape":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var imageShadowEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                leaf = DirectChildRange(xml, imageShadowEffectList, "effectLst", "outerShdw", operation);
                attribute = "rotWithShape";
                break;
            case "imageShadowBlurRadiusEmu":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var imageShadowBlurEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                leaf = DirectChildRange(xml, imageShadowBlurEffectList, "effectLst", "outerShdw", operation);
                attribute = "blurRad";
                break;
            case "imageShadowDistanceEmu":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var imageShadowDistanceEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                leaf = DirectChildRange(xml, imageShadowDistanceEffectList, "effectLst", "outerShdw", operation);
                attribute = "dist";
                break;
            case "imageShadowDirectionDegrees":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var imageShadowDirectionEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                leaf = DirectChildRange(xml, imageShadowDirectionEffectList, "effectLst", "outerShdw", operation);
                attribute = "dir";
                break;
            case "imageShadowAlignment":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var imageShadowAlignmentEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                leaf = DirectChildRange(xml, imageShadowAlignmentEffectList, "effectLst", "outerShdw", operation);
                attribute = "algn";
                break;
            case "imageShadowOpacityThousandthPercent":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose image shadow opacity.", operation.SlidePartPath);
                var imageShadowOpacityEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                var imageShadowOpacityOuter = DirectChildRange(xml, imageShadowOpacityEffectList, "effectLst", "outerShdw", operation);
                var imageShadowOpacityColor = DirectChildRanges(xml, imageShadowOpacityOuter)
                    .Where(entry => entry.LocalName is "srgbClr" or "schemeClr")
                    .ToArray();
                if (imageShadowOpacityColor.Length != 1)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} image shadow color is missing or ambiguous.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, imageShadowOpacityColor[0], imageShadowOpacityColor[0].LocalName, "alpha", operation);
                attribute = "val";
                break;
            case "shadowColorRgb":
            case "shadowColorScheme":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var shadowColorEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                var shadowOuter = DirectChildRange(xml, shadowColorEffectList, "effectLst", "outerShdw", operation);
                var shadowColorName = LeafKind(operation) == "shadowColorRgb" ? "srgbClr" : "schemeClr";
                leaf = DirectChildRange(xml, shadowOuter, "outerShdw", shadowColorName, operation);
                attribute = "val";
                break;
            case "imageShadowColorRgb":
            case "imageShadowColorScheme":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose image shadow color.", operation.SlidePartPath);
                var imageShadowColorEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                var imageShadowColorOuter = DirectChildRange(xml, imageShadowColorEffectList, "effectLst", "outerShdw", operation);
                var imageShadowColorName = LeafKind(operation) == "imageShadowColorRgb" ? "srgbClr" : "schemeClr";
                leaf = DirectChildRange(xml, imageShadowColorOuter, "outerShdw", imageShadowColorName, operation);
                attribute = "val";
                break;
            case "shadowOpacityThousandthPercent":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose shadow opacity.", operation.SlidePartPath);
                var shadowOpacityEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                var shadowOpacityOuter = DirectChildRange(xml, shadowOpacityEffectList, "effectLst", "outerShdw", operation);
                var shadowOpacityColor = DirectChildRanges(xml, shadowOpacityOuter)
                    .Where(entry => entry.LocalName is "srgbClr" or "schemeClr")
                    .ToArray();
                if (shadowOpacityColor.Length != 1)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} shadow color is missing or ambiguous.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, shadowOpacityColor[0], shadowOpacityColor[0].LocalName, "alpha", operation);
                attribute = "val";
                break;
            case "imageOpacityThousandthPercent":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose image opacity.", operation.SlidePartPath);
                var blipFill = DirectChildRange(xml, elementRange, owner, "blipFill", operation);
                var blip = DirectChildRange(xml, blipFill, "blipFill", "blip", operation);
                leaf = DirectChildRange(xml, blip, "blip", "alphaModFix", operation);
                attribute = "amt";
                break;
            case "imageMaskPreset":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose an image mask preset.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "prstGeom", operation);
                attribute = "prst";
                break;
            case "imageMaskAdjustment":
                if (owner != "pic") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose an image mask adjustment.", operation.SlidePartPath);
                var maskGeometry = DirectChildRange(xml, properties, "spPr", "prstGeom", operation);
                var maskAdjustmentList = DirectChildRange(xml, maskGeometry, "prstGeom", "avLst", operation);
                var maskAdjustmentGuides = DirectChildRanges(xml, maskAdjustmentList)
                    .Where(entry => entry.LocalName == "gd")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)maskAdjustmentGuides.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} image mask adjustment index is out of range.", operation.SlidePartPath);
                leaf = maskAdjustmentGuides[(int)operation.NativeLeafIndex];
                attribute = "fmla";
                break;
            case "customGeometryAdjustment":
            case "customGeometryAdjustmentFormula":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry adjustment.", operation.SlidePartPath);
                var customGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var adjustmentList = DirectChildRange(xml, customGeometry, "custGeom", "avLst", operation);
                var adjustmentGuides = DirectChildRanges(xml, adjustmentList)
                    .Where(entry => entry.LocalName == "gd")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)adjustmentGuides.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment index is out of range.", operation.SlidePartPath);
                leaf = adjustmentGuides[(int)operation.NativeLeafIndex];
                attribute = "fmla";
                break;
            case "customGeometryPathWidth":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry path width.", operation.SlidePartPath);
                var customPathWidthGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathWidthList = DirectChildRange(xml, customPathWidthGeometry, "custGeom", "pathLst", operation);
                var customPathWidthEntries = DirectChildRanges(xml, customPathWidthList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathWidthEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                leaf = customPathWidthEntries[(int)operation.NativeLeafIndex];
                attribute = "w";
                break;
            case "customGeometryPathHeight":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry path height.", operation.SlidePartPath);
                var customPathHeightGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathHeightList = DirectChildRange(xml, customPathHeightGeometry, "custGeom", "pathLst", operation);
                var customPathHeightEntries = DirectChildRanges(xml, customPathHeightList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathHeightEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                leaf = customPathHeightEntries[(int)operation.NativeLeafIndex];
                attribute = "h";
                break;
            case "customGeometryPathArcWidthRadius":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry arc width radius.", operation.SlidePartPath);
                var customPathArcWidthRadiusGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathArcWidthRadiusList = DirectChildRange(xml, customPathArcWidthRadiusGeometry, "custGeom", "pathLst", operation);
                var customPathArcWidthRadiusEntries = DirectChildRanges(xml, customPathArcWidthRadiusList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathArcWidthRadiusEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathArcWidthRadiusCommands = DirectChildRanges(xml, customPathArcWidthRadiusEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathArcWidthRadiusCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathArcWidthRadiusCommand = customPathArcWidthRadiusCommands[(int)operation.TextLeafIndex];
                if (customPathArcWidthRadiusCommand.LocalName != "arcTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an arc-to command.", operation.SlidePartPath);
                leaf = customPathArcWidthRadiusCommand;
                attribute = "wR";
                break;
            case "customGeometryPathArcHeightRadius":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry arc height radius.", operation.SlidePartPath);
                var customPathArcHeightRadiusGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathArcHeightRadiusList = DirectChildRange(xml, customPathArcHeightRadiusGeometry, "custGeom", "pathLst", operation);
                var customPathArcHeightRadiusEntries = DirectChildRanges(xml, customPathArcHeightRadiusList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathArcHeightRadiusEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathArcHeightRadiusCommands = DirectChildRanges(xml, customPathArcHeightRadiusEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathArcHeightRadiusCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathArcHeightRadiusCommand = customPathArcHeightRadiusCommands[(int)operation.TextLeafIndex];
                if (customPathArcHeightRadiusCommand.LocalName != "arcTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an arc-to command.", operation.SlidePartPath);
                leaf = customPathArcHeightRadiusCommand;
                attribute = "hR";
                break;
            case "customGeometryPathArcStartAngle60000":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry arc start angle.", operation.SlidePartPath);
                var customPathArcStartAngleGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathArcStartAngleList = DirectChildRange(xml, customPathArcStartAngleGeometry, "custGeom", "pathLst", operation);
                var customPathArcStartAngleEntries = DirectChildRanges(xml, customPathArcStartAngleList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathArcStartAngleEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathArcStartAngleCommands = DirectChildRanges(xml, customPathArcStartAngleEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathArcStartAngleCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathArcStartAngleCommand = customPathArcStartAngleCommands[(int)operation.TextLeafIndex];
                if (customPathArcStartAngleCommand.LocalName != "arcTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an arc-to command.", operation.SlidePartPath);
                leaf = customPathArcStartAngleCommand;
                attribute = "stAng";
                break;
            case "customGeometryPathArcSweepAngle60000":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry arc sweep angle.", operation.SlidePartPath);
                var customPathArcSweepAngleGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathArcSweepAngleList = DirectChildRange(xml, customPathArcSweepAngleGeometry, "custGeom", "pathLst", operation);
                var customPathArcSweepAngleEntries = DirectChildRanges(xml, customPathArcSweepAngleList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathArcSweepAngleEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathArcSweepAngleCommands = DirectChildRanges(xml, customPathArcSweepAngleEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathArcSweepAngleCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathArcSweepAngleCommand = customPathArcSweepAngleCommands[(int)operation.TextLeafIndex];
                if (customPathArcSweepAngleCommand.LocalName != "arcTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an arc-to command.", operation.SlidePartPath);
                leaf = customPathArcSweepAngleCommand;
                attribute = "swAng";
                break;
            case "customGeometryPathLineToX":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry line-to x coordinate.", operation.SlidePartPath);
                var customPathLineToXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathLineToXList = DirectChildRange(xml, customPathLineToXGeometry, "custGeom", "pathLst", operation);
                var customPathLineToXEntries = DirectChildRanges(xml, customPathLineToXList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathLineToXEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathLineToXCommands = DirectChildRanges(xml, customPathLineToXEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathLineToXCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathLineToXCommand = customPathLineToXCommands[(int)operation.TextLeafIndex];
                if (customPathLineToXCommand.LocalName != "lnTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a line-to command.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, customPathLineToXCommand, "lnTo", "pt", operation);
                attribute = "x";
                break;
            case "customGeometryPathLineToY":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry line-to y coordinate.", operation.SlidePartPath);
                var customPathLineToYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathLineToYList = DirectChildRange(xml, customPathLineToYGeometry, "custGeom", "pathLst", operation);
                var customPathLineToYEntries = DirectChildRanges(xml, customPathLineToYList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathLineToYEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathLineToYCommands = DirectChildRanges(xml, customPathLineToYEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathLineToYCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathLineToYCommand = customPathLineToYCommands[(int)operation.TextLeafIndex];
                if (customPathLineToYCommand.LocalName != "lnTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a line-to command.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, customPathLineToYCommand, "lnTo", "pt", operation);
                attribute = "y";
                break;
            case "customGeometryPathMoveToX":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry move-to x coordinate.", operation.SlidePartPath);
                var customPathMoveToXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathMoveToXList = DirectChildRange(xml, customPathMoveToXGeometry, "custGeom", "pathLst", operation);
                var customPathMoveToXEntries = DirectChildRanges(xml, customPathMoveToXList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathMoveToXEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathMoveToXCommands = DirectChildRanges(xml, customPathMoveToXEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathMoveToXCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathMoveToXCommand = customPathMoveToXCommands[(int)operation.TextLeafIndex];
                if (customPathMoveToXCommand.LocalName != "moveTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a move-to command.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, customPathMoveToXCommand, "moveTo", "pt", operation);
                attribute = "x";
                break;
            case "customGeometryPathMoveToY":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry move-to y coordinate.", operation.SlidePartPath);
                var customPathMoveToYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathMoveToYList = DirectChildRange(xml, customPathMoveToYGeometry, "custGeom", "pathLst", operation);
                var customPathMoveToYEntries = DirectChildRanges(xml, customPathMoveToYList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathMoveToYEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathMoveToYCommands = DirectChildRanges(xml, customPathMoveToYEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathMoveToYCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathMoveToYCommand = customPathMoveToYCommands[(int)operation.TextLeafIndex];
                if (customPathMoveToYCommand.LocalName != "moveTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a move-to command.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, customPathMoveToYCommand, "moveTo", "pt", operation);
                attribute = "y";
                break;
            case "customGeometryPathQuadraticEndX":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry quadratic end-point x coordinate.", operation.SlidePartPath);
                var customPathQuadraticEndXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathQuadraticEndXList = DirectChildRange(xml, customPathQuadraticEndXGeometry, "custGeom", "pathLst", operation);
                var customPathQuadraticEndXEntries = DirectChildRanges(xml, customPathQuadraticEndXList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathQuadraticEndXEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathQuadraticEndXCommands = DirectChildRanges(xml, customPathQuadraticEndXEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathQuadraticEndXCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathQuadraticEndXCommand = customPathQuadraticEndXCommands[(int)operation.TextLeafIndex];
                if (customPathQuadraticEndXCommand.LocalName != "quadBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a quadratic-bezier command.", operation.SlidePartPath);
                var customPathQuadraticEndXPoints = DirectChildRanges(xml, customPathQuadraticEndXCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathQuadraticEndXPoints.Length != 2)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} quadratic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathQuadraticEndXPoints[1];
                attribute = "x";
                break;
            case "customGeometryPathQuadraticEndY":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry quadratic end-point y coordinate.", operation.SlidePartPath);
                var customPathQuadraticEndYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathQuadraticEndYList = DirectChildRange(xml, customPathQuadraticEndYGeometry, "custGeom", "pathLst", operation);
                var customPathQuadraticEndYEntries = DirectChildRanges(xml, customPathQuadraticEndYList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathQuadraticEndYEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathQuadraticEndYCommands = DirectChildRanges(xml, customPathQuadraticEndYEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathQuadraticEndYCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathQuadraticEndYCommand = customPathQuadraticEndYCommands[(int)operation.TextLeafIndex];
                if (customPathQuadraticEndYCommand.LocalName != "quadBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a quadratic-bezier command.", operation.SlidePartPath);
                var customPathQuadraticEndYPoints = DirectChildRanges(xml, customPathQuadraticEndYCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathQuadraticEndYPoints.Length != 2)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} quadratic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathQuadraticEndYPoints[1];
                attribute = "y";
                break;
            case "customGeometryPathQuadraticControlX":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry quadratic control-point x coordinate.", operation.SlidePartPath);
                var customPathQuadraticControlXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathQuadraticControlXList = DirectChildRange(xml, customPathQuadraticControlXGeometry, "custGeom", "pathLst", operation);
                var customPathQuadraticControlXEntries = DirectChildRanges(xml, customPathQuadraticControlXList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathQuadraticControlXEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathQuadraticControlXCommands = DirectChildRanges(xml, customPathQuadraticControlXEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathQuadraticControlXCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathQuadraticControlXCommand = customPathQuadraticControlXCommands[(int)operation.TextLeafIndex];
                if (customPathQuadraticControlXCommand.LocalName != "quadBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a quadratic-bezier command.", operation.SlidePartPath);
                var customPathQuadraticControlXPoints = DirectChildRanges(xml, customPathQuadraticControlXCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathQuadraticControlXPoints.Length != 2)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} quadratic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathQuadraticControlXPoints[0];
                attribute = "x";
                break;
            case "customGeometryPathQuadraticControlY":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry quadratic control-point y coordinate.", operation.SlidePartPath);
                var customPathQuadraticControlYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathQuadraticControlYList = DirectChildRange(xml, customPathQuadraticControlYGeometry, "custGeom", "pathLst", operation);
                var customPathQuadraticControlYEntries = DirectChildRanges(xml, customPathQuadraticControlYList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathQuadraticControlYEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathQuadraticControlYCommands = DirectChildRanges(xml, customPathQuadraticControlYEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathQuadraticControlYCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathQuadraticControlYCommand = customPathQuadraticControlYCommands[(int)operation.TextLeafIndex];
                if (customPathQuadraticControlYCommand.LocalName != "quadBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a quadratic-bezier command.", operation.SlidePartPath);
                var customPathQuadraticControlYPoints = DirectChildRanges(xml, customPathQuadraticControlYCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathQuadraticControlYPoints.Length != 2)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} quadratic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathQuadraticControlYPoints[0];
                attribute = "y";
                break;
            case "customGeometryPathCubicEndX":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry cubic end-point x coordinate.", operation.SlidePartPath);
                var customPathCubicEndXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathCubicEndXList = DirectChildRange(xml, customPathCubicEndXGeometry, "custGeom", "pathLst", operation);
                var customPathCubicEndXEntries = DirectChildRanges(xml, customPathCubicEndXList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathCubicEndXEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathCubicEndXCommands = DirectChildRanges(xml, customPathCubicEndXEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathCubicEndXCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathCubicEndXCommand = customPathCubicEndXCommands[(int)operation.TextLeafIndex];
                if (customPathCubicEndXCommand.LocalName != "cubicBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a cubic-bezier command.", operation.SlidePartPath);
                var customPathCubicEndXPoints = DirectChildRanges(xml, customPathCubicEndXCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathCubicEndXPoints.Length != 3)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} cubic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathCubicEndXPoints[2];
                attribute = "x";
                break;
            case "customGeometryPathCubicEndY":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry cubic end-point y coordinate.", operation.SlidePartPath);
                var customPathCubicEndYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathCubicEndYList = DirectChildRange(xml, customPathCubicEndYGeometry, "custGeom", "pathLst", operation);
                var customPathCubicEndYEntries = DirectChildRanges(xml, customPathCubicEndYList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathCubicEndYEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathCubicEndYCommands = DirectChildRanges(xml, customPathCubicEndYEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathCubicEndYCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathCubicEndYCommand = customPathCubicEndYCommands[(int)operation.TextLeafIndex];
                if (customPathCubicEndYCommand.LocalName != "cubicBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a cubic-bezier command.", operation.SlidePartPath);
                var customPathCubicEndYPoints = DirectChildRanges(xml, customPathCubicEndYCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathCubicEndYPoints.Length != 3)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} cubic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathCubicEndYPoints[2];
                attribute = "y";
                break;
            case "customGeometryPathCubicControl1X":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry cubic first-control-point x coordinate.", operation.SlidePartPath);
                var customPathCubicControl1XGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathCubicControl1XList = DirectChildRange(xml, customPathCubicControl1XGeometry, "custGeom", "pathLst", operation);
                var customPathCubicControl1XEntries = DirectChildRanges(xml, customPathCubicControl1XList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathCubicControl1XEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathCubicControl1XCommands = DirectChildRanges(xml, customPathCubicControl1XEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathCubicControl1XCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathCubicControl1XCommand = customPathCubicControl1XCommands[(int)operation.TextLeafIndex];
                if (customPathCubicControl1XCommand.LocalName != "cubicBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a cubic-bezier command.", operation.SlidePartPath);
                var customPathCubicControl1XPoints = DirectChildRanges(xml, customPathCubicControl1XCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathCubicControl1XPoints.Length != 3)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} cubic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathCubicControl1XPoints[0];
                attribute = "x";
                break;
            case "customGeometryPathCubicControl1Y":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry cubic first-control-point y coordinate.", operation.SlidePartPath);
                var customPathCubicControl1YGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathCubicControl1YList = DirectChildRange(xml, customPathCubicControl1YGeometry, "custGeom", "pathLst", operation);
                var customPathCubicControl1YEntries = DirectChildRanges(xml, customPathCubicControl1YList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathCubicControl1YEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathCubicControl1YCommands = DirectChildRanges(xml, customPathCubicControl1YEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathCubicControl1YCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathCubicControl1YCommand = customPathCubicControl1YCommands[(int)operation.TextLeafIndex];
                if (customPathCubicControl1YCommand.LocalName != "cubicBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a cubic-bezier command.", operation.SlidePartPath);
                var customPathCubicControl1YPoints = DirectChildRanges(xml, customPathCubicControl1YCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathCubicControl1YPoints.Length != 3)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} cubic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathCubicControl1YPoints[0];
                attribute = "y";
                break;
            case "customGeometryPathCubicControl2X":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry cubic second-control-point x coordinate.", operation.SlidePartPath);
                var customPathCubicControl2XGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathCubicControl2XList = DirectChildRange(xml, customPathCubicControl2XGeometry, "custGeom", "pathLst", operation);
                var customPathCubicControl2XEntries = DirectChildRanges(xml, customPathCubicControl2XList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathCubicControl2XEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathCubicControl2XCommands = DirectChildRanges(xml, customPathCubicControl2XEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathCubicControl2XCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathCubicControl2XCommand = customPathCubicControl2XCommands[(int)operation.TextLeafIndex];
                if (customPathCubicControl2XCommand.LocalName != "cubicBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a cubic-bezier command.", operation.SlidePartPath);
                var customPathCubicControl2XPoints = DirectChildRanges(xml, customPathCubicControl2XCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathCubicControl2XPoints.Length != 3)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} cubic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathCubicControl2XPoints[1];
                attribute = "x";
                break;
            case "customGeometryPathCubicControl2Y":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry cubic second-control-point y coordinate.", operation.SlidePartPath);
                var customPathCubicControl2YGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathCubicControl2YList = DirectChildRange(xml, customPathCubicControl2YGeometry, "custGeom", "pathLst", operation);
                var customPathCubicControl2YEntries = DirectChildRanges(xml, customPathCubicControl2YList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathCubicControl2YEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                var customPathCubicControl2YCommands = DirectChildRanges(xml, customPathCubicControl2YEntries[(int)operation.NativeLeafIndex]);
                if (operation.TextLeafIndex >= (uint)customPathCubicControl2YCommands.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry command index is out of range.", operation.SlidePartPath);
                var customPathCubicControl2YCommand = customPathCubicControl2YCommands[(int)operation.TextLeafIndex];
                if (customPathCubicControl2YCommand.LocalName != "cubicBezTo")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a cubic-bezier command.", operation.SlidePartPath);
                var customPathCubicControl2YPoints = DirectChildRanges(xml, customPathCubicControl2YCommand)
                    .Where(entry => entry.LocalName == "pt")
                    .ToArray();
                if (customPathCubicControl2YPoints.Length != 3)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} cubic-bezier point topology is not canonical.", operation.SlidePartPath);
                leaf = customPathCubicControl2YPoints[1];
                attribute = "y";
                break;
            case "customGeometryPathFill":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry path fill.", operation.SlidePartPath);
                var customPathFillGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathFillList = DirectChildRange(xml, customPathFillGeometry, "custGeom", "pathLst", operation);
                var customPathFillEntries = DirectChildRanges(xml, customPathFillList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathFillEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                leaf = customPathFillEntries[(int)operation.NativeLeafIndex];
                attribute = "fill";
                break;
            case "customGeometryPathStroke":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry path stroke.", operation.SlidePartPath);
                var customPathStrokeGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathStrokeList = DirectChildRange(xml, customPathStrokeGeometry, "custGeom", "pathLst", operation);
                var customPathStrokeEntries = DirectChildRanges(xml, customPathStrokeList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathStrokeEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                leaf = customPathStrokeEntries[(int)operation.NativeLeafIndex];
                attribute = "stroke";
                break;
            case "customGeometryPathExtrusionAllowed":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry path extrusion permission.", operation.SlidePartPath);
                var customPathExtrusionGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPathExtrusionList = DirectChildRange(xml, customPathExtrusionGeometry, "custGeom", "pathLst", operation);
                var customPathExtrusionEntries = DirectChildRanges(xml, customPathExtrusionList)
                    .Where(entry => entry.LocalName == "path")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customPathExtrusionEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry path index is out of range.", operation.SlidePartPath);
                leaf = customPathExtrusionEntries[(int)operation.NativeLeafIndex];
                attribute = "extrusionOk";
                break;
            case "customGeometryGuide":
            case "customGeometryGuideFormula":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry guide.", operation.SlidePartPath);
                var customGuideGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customGuideList = DirectChildRange(xml, customGuideGeometry, "custGeom", "gdLst", operation);
                var customGuideEntries = DirectChildRanges(xml, customGuideList)
                    .Where(entry => entry.LocalName == "gd")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customGuideEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry guide index is out of range.", operation.SlidePartPath);
                leaf = customGuideEntries[(int)operation.NativeLeafIndex];
                attribute = "fmla";
                break;
            case "customGeometryConnectionSiteAngle60000":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry connection-site angle.", operation.SlidePartPath);
                var customConnectionGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customConnectionList = DirectChildRange(xml, customConnectionGeometry, "custGeom", "cxnLst", operation);
                var customConnectionEntries = DirectChildRanges(xml, customConnectionList)
                    .Where(entry => entry.LocalName == "cxn")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customConnectionEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry connection-site index is out of range.", operation.SlidePartPath);
                leaf = customConnectionEntries[(int)operation.NativeLeafIndex];
                attribute = "ang";
                break;
            case "customGeometryConnectionSiteXEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry connection-site x coordinate.", operation.SlidePartPath);
                var customConnectionXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customConnectionXList = DirectChildRange(xml, customConnectionXGeometry, "custGeom", "cxnLst", operation);
                var customConnectionXEntries = DirectChildRanges(xml, customConnectionXList)
                    .Where(entry => entry.LocalName == "cxn")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customConnectionXEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry connection-site index is out of range.", operation.SlidePartPath);
                var customConnectionXSite = customConnectionXEntries[(int)operation.NativeLeafIndex];
                var connectionXPosition = DirectChildRange(xml, customConnectionXSite, "cxn", "pos", operation);
                leaf = connectionXPosition;
                attribute = "x";
                break;
            case "customGeometryConnectionSiteYEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry connection-site y coordinate.", operation.SlidePartPath);
                var customConnectionYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customConnectionYList = DirectChildRange(xml, customConnectionYGeometry, "custGeom", "cxnLst", operation);
                var customConnectionYEntries = DirectChildRanges(xml, customConnectionYList)
                    .Where(entry => entry.LocalName == "cxn")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)customConnectionYEntries.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry connection-site index is out of range.", operation.SlidePartPath);
                var customConnectionYSite = customConnectionYEntries[(int)operation.NativeLeafIndex];
                var connectionYPosition = DirectChildRange(xml, customConnectionYSite, "cxn", "pos", operation);
                leaf = connectionYPosition;
                attribute = "y";
                break;
            case "customGeometryAdjustmentHandleXEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry XY adjustment-handle x position.", operation.SlidePartPath);
                var customHandleXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customHandleXList = DirectChildRange(xml, customHandleXGeometry, "custGeom", "ahLst", operation);
                var customHandleXEntries = DirectChildRanges(xml, customHandleXList);
                if (operation.NativeLeafIndex >= (uint)customHandleXEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customHandleX = customHandleXEntries[(int)operation.NativeLeafIndex];
                if (customHandleX.LocalName != "ahXY")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an XY adjustment handle.", operation.SlidePartPath);
                var handleXPosition = DirectChildRange(xml, customHandleX, "ahXY", "pos", operation);
                leaf = handleXPosition;
                attribute = "x";
                break;
            case "customGeometryAdjustmentHandleYEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry XY adjustment-handle y position.", operation.SlidePartPath);
                var customHandleYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customHandleYList = DirectChildRange(xml, customHandleYGeometry, "custGeom", "ahLst", operation);
                var customHandleYEntries = DirectChildRanges(xml, customHandleYList);
                if (operation.NativeLeafIndex >= (uint)customHandleYEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customHandleY = customHandleYEntries[(int)operation.NativeLeafIndex];
                if (customHandleY.LocalName != "ahXY")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an XY adjustment handle.", operation.SlidePartPath);
                var handleYPosition = DirectChildRange(xml, customHandleY, "ahXY", "pos", operation);
                leaf = handleYPosition;
                attribute = "y";
                break;
            case "customGeometryAdjustmentHandleMinXEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry XY adjustment-handle minimum-x bound.", operation.SlidePartPath);
                var customHandleMinXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customHandleMinXList = DirectChildRange(xml, customHandleMinXGeometry, "custGeom", "ahLst", operation);
                var customHandleMinXEntries = DirectChildRanges(xml, customHandleMinXList);
                if (operation.NativeLeafIndex >= (uint)customHandleMinXEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customHandleMinX = customHandleMinXEntries[(int)operation.NativeLeafIndex];
                if (customHandleMinX.LocalName != "ahXY")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an XY adjustment handle.", operation.SlidePartPath);
                leaf = customHandleMinX;
                attribute = "minX";
                break;
            case "customGeometryAdjustmentHandleMaxXEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry XY adjustment-handle maximum-x bound.", operation.SlidePartPath);
                var customHandleMaxXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customHandleMaxXList = DirectChildRange(xml, customHandleMaxXGeometry, "custGeom", "ahLst", operation);
                var customHandleMaxXEntries = DirectChildRanges(xml, customHandleMaxXList);
                if (operation.NativeLeafIndex >= (uint)customHandleMaxXEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customHandleMaxX = customHandleMaxXEntries[(int)operation.NativeLeafIndex];
                if (customHandleMaxX.LocalName != "ahXY")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an XY adjustment handle.", operation.SlidePartPath);
                leaf = customHandleMaxX;
                attribute = "maxX";
                break;
            case "customGeometryAdjustmentHandleMinYEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry XY adjustment-handle minimum-y bound.", operation.SlidePartPath);
                var customHandleMinYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customHandleMinYList = DirectChildRange(xml, customHandleMinYGeometry, "custGeom", "ahLst", operation);
                var customHandleMinYEntries = DirectChildRanges(xml, customHandleMinYList);
                if (operation.NativeLeafIndex >= (uint)customHandleMinYEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customHandleMinY = customHandleMinYEntries[(int)operation.NativeLeafIndex];
                if (customHandleMinY.LocalName != "ahXY")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an XY adjustment handle.", operation.SlidePartPath);
                leaf = customHandleMinY;
                attribute = "minY";
                break;
            case "customGeometryAdjustmentHandleMaxYEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry XY adjustment-handle maximum-y bound.", operation.SlidePartPath);
                var customHandleMaxYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customHandleMaxYList = DirectChildRange(xml, customHandleMaxYGeometry, "custGeom", "ahLst", operation);
                var customHandleMaxYEntries = DirectChildRanges(xml, customHandleMaxYList);
                if (operation.NativeLeafIndex >= (uint)customHandleMaxYEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customHandleMaxY = customHandleMaxYEntries[(int)operation.NativeLeafIndex];
                if (customHandleMaxY.LocalName != "ahXY")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not an XY adjustment handle.", operation.SlidePartPath);
                leaf = customHandleMaxY;
                attribute = "maxY";
                break;
            case "customGeometryAdjustmentHandlePolarMinRadiusEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry polar adjustment-handle minimum-radius bound.", operation.SlidePartPath);
                var customPolarHandleMinRadiusGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPolarHandleMinRadiusList = DirectChildRange(xml, customPolarHandleMinRadiusGeometry, "custGeom", "ahLst", operation);
                var customPolarHandleMinRadiusEntries = DirectChildRanges(xml, customPolarHandleMinRadiusList);
                if (operation.NativeLeafIndex >= (uint)customPolarHandleMinRadiusEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customPolarHandleMinRadius = customPolarHandleMinRadiusEntries[(int)operation.NativeLeafIndex];
                if (customPolarHandleMinRadius.LocalName != "ahPolar")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a polar adjustment handle.", operation.SlidePartPath);
                leaf = customPolarHandleMinRadius;
                attribute = "minR";
                break;
            case "customGeometryAdjustmentHandlePolarMaxRadiusEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry polar adjustment-handle maximum-radius bound.", operation.SlidePartPath);
                var customPolarHandleMaxRadiusGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPolarHandleMaxRadiusList = DirectChildRange(xml, customPolarHandleMaxRadiusGeometry, "custGeom", "ahLst", operation);
                var customPolarHandleMaxRadiusEntries = DirectChildRanges(xml, customPolarHandleMaxRadiusList);
                if (operation.NativeLeafIndex >= (uint)customPolarHandleMaxRadiusEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customPolarHandleMaxRadius = customPolarHandleMaxRadiusEntries[(int)operation.NativeLeafIndex];
                if (customPolarHandleMaxRadius.LocalName != "ahPolar")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a polar adjustment handle.", operation.SlidePartPath);
                leaf = customPolarHandleMaxRadius;
                attribute = "maxR";
                break;
            case "customGeometryAdjustmentHandlePolarMinAngle60000":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry polar adjustment-handle minimum-angle bound.", operation.SlidePartPath);
                var customPolarHandleMinAngleGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPolarHandleMinAngleList = DirectChildRange(xml, customPolarHandleMinAngleGeometry, "custGeom", "ahLst", operation);
                var customPolarHandleMinAngleEntries = DirectChildRanges(xml, customPolarHandleMinAngleList);
                if (operation.NativeLeafIndex >= (uint)customPolarHandleMinAngleEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customPolarHandleMinAngle = customPolarHandleMinAngleEntries[(int)operation.NativeLeafIndex];
                if (customPolarHandleMinAngle.LocalName != "ahPolar")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a polar adjustment handle.", operation.SlidePartPath);
                leaf = customPolarHandleMinAngle;
                attribute = "minAng";
                break;
            case "customGeometryAdjustmentHandlePolarMaxAngle60000":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry polar adjustment-handle maximum-angle bound.", operation.SlidePartPath);
                var customPolarHandleMaxAngleGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPolarHandleMaxAngleList = DirectChildRange(xml, customPolarHandleMaxAngleGeometry, "custGeom", "ahLst", operation);
                var customPolarHandleMaxAngleEntries = DirectChildRanges(xml, customPolarHandleMaxAngleList);
                if (operation.NativeLeafIndex >= (uint)customPolarHandleMaxAngleEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customPolarHandleMaxAngle = customPolarHandleMaxAngleEntries[(int)operation.NativeLeafIndex];
                if (customPolarHandleMaxAngle.LocalName != "ahPolar")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a polar adjustment handle.", operation.SlidePartPath);
                leaf = customPolarHandleMaxAngle;
                attribute = "maxAng";
                break;
            case "customGeometryAdjustmentHandlePolarXEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry polar adjustment-handle x position.", operation.SlidePartPath);
                var customPolarHandleXGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPolarHandleXList = DirectChildRange(xml, customPolarHandleXGeometry, "custGeom", "ahLst", operation);
                var customPolarHandleXEntries = DirectChildRanges(xml, customPolarHandleXList);
                if (operation.NativeLeafIndex >= (uint)customPolarHandleXEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customPolarHandleX = customPolarHandleXEntries[(int)operation.NativeLeafIndex];
                if (customPolarHandleX.LocalName != "ahPolar")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a polar adjustment handle.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, customPolarHandleX, "ahPolar", "pos", operation);
                attribute = "x";
                break;
            case "customGeometryAdjustmentHandlePolarYEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry polar adjustment-handle y position.", operation.SlidePartPath);
                var customPolarHandleYGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customPolarHandleYList = DirectChildRange(xml, customPolarHandleYGeometry, "custGeom", "ahLst", operation);
                var customPolarHandleYEntries = DirectChildRanges(xml, customPolarHandleYList);
                if (operation.NativeLeafIndex >= (uint)customPolarHandleYEntries.Count)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry adjustment-handle index is out of range.", operation.SlidePartPath);
                var customPolarHandleY = customPolarHandleYEntries[(int)operation.NativeLeafIndex];
                if (customPolarHandleY.LocalName != "ahPolar")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a polar adjustment handle.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, customPolarHandleY, "ahPolar", "pos", operation);
                attribute = "y";
                break;
            case "customGeometryTextRectangleLeftEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry text-rectangle left edge.", operation.SlidePartPath);
                var customTextRectangleGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customTextRectangle = DirectChildRange(xml, customTextRectangleGeometry, "custGeom", "rect", operation);
                if (!TryXmlAttributeValue(xml, customTextRectangle, "l", out var customTextRectangleLeftToken))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle has no left edge.", operation.SlidePartPath);
                if (customTextRectangleLeftToken == "officeKitTextLeft")
                {
                    var customTextRectangleGuides = DirectChildRange(xml, customTextRectangleGeometry, "custGeom", "gdLst", operation);
                    var customTextRectangleGuideEntries = DirectChildRanges(xml, customTextRectangleGuides)
                        .Where(entry => entry.LocalName == "gd" && TryXmlAttributeValue(xml, entry, "name", out var name) && name == "officeKitTextLeft")
                        .ToArray();
                    if (customTextRectangleGuideEntries.Length != 1)
                        throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle private left guide is missing or ambiguous.", operation.SlidePartPath);
                    leaf = customTextRectangleGuideEntries[0];
                    attribute = "fmla";
                }
                else
                {
                    leaf = customTextRectangle;
                    attribute = "l";
                }
                break;
            case "customGeometryTextRectangleTopEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry text-rectangle top edge.", operation.SlidePartPath);
                var customTextRectangleTopGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customTextRectangleTop = DirectChildRange(xml, customTextRectangleTopGeometry, "custGeom", "rect", operation);
                if (!TryXmlAttributeValue(xml, customTextRectangleTop, "t", out var customTextRectangleTopToken))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle has no top edge.", operation.SlidePartPath);
                if (customTextRectangleTopToken == "officeKitTextTop")
                {
                    var customTextRectangleTopGuides = DirectChildRange(xml, customTextRectangleTopGeometry, "custGeom", "gdLst", operation);
                    var customTextRectangleTopGuideEntries = DirectChildRanges(xml, customTextRectangleTopGuides)
                        .Where(entry => entry.LocalName == "gd" && TryXmlAttributeValue(xml, entry, "name", out var name) && name == "officeKitTextTop")
                        .ToArray();
                    if (customTextRectangleTopGuideEntries.Length != 1)
                        throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle private top guide is missing or ambiguous.", operation.SlidePartPath);
                    leaf = customTextRectangleTopGuideEntries[0];
                    attribute = "fmla";
                }
                else
                {
                    leaf = customTextRectangleTop;
                    attribute = "t";
                }
                break;
            case "customGeometryTextRectangleRightEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry text-rectangle right edge.", operation.SlidePartPath);
                var customTextRectangleRightGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customTextRectangleRight = DirectChildRange(xml, customTextRectangleRightGeometry, "custGeom", "rect", operation);
                if (!TryXmlAttributeValue(xml, customTextRectangleRight, "r", out var customTextRectangleRightToken))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle has no right edge.", operation.SlidePartPath);
                if (customTextRectangleRightToken == "officeKitTextRight")
                {
                    var customTextRectangleRightGuides = DirectChildRange(xml, customTextRectangleRightGeometry, "custGeom", "gdLst", operation);
                    var customTextRectangleRightGuideEntries = DirectChildRanges(xml, customTextRectangleRightGuides)
                        .Where(entry => entry.LocalName == "gd" && TryXmlAttributeValue(xml, entry, "name", out var name) && name == "officeKitTextRight")
                        .ToArray();
                    if (customTextRectangleRightGuideEntries.Length != 1)
                        throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle private right guide is missing or ambiguous.", operation.SlidePartPath);
                    leaf = customTextRectangleRightGuideEntries[0];
                    attribute = "fmla";
                }
                else
                {
                    leaf = customTextRectangleRight;
                    attribute = "r";
                }
                break;
            case "customGeometryTextRectangleBottomEmu":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a custom-geometry text-rectangle bottom edge.", operation.SlidePartPath);
                var customTextRectangleBottomGeometry = DirectChildRange(xml, properties, "spPr", "custGeom", operation);
                var customTextRectangleBottom = DirectChildRange(xml, customTextRectangleBottomGeometry, "custGeom", "rect", operation);
                if (!TryXmlAttributeValue(xml, customTextRectangleBottom, "b", out var customTextRectangleBottomToken))
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle has no bottom edge.", operation.SlidePartPath);
                if (customTextRectangleBottomToken == "officeKitTextBottom")
                {
                    var customTextRectangleBottomGuides = DirectChildRange(xml, customTextRectangleBottomGeometry, "custGeom", "gdLst", operation);
                    var customTextRectangleBottomGuideEntries = DirectChildRanges(xml, customTextRectangleBottomGuides)
                        .Where(entry => entry.LocalName == "gd" && TryXmlAttributeValue(xml, entry, "name", out var name) && name == "officeKitTextBottom")
                        .ToArray();
                    if (customTextRectangleBottomGuideEntries.Length != 1)
                        throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} custom-geometry text rectangle private bottom guide is missing or ambiguous.", operation.SlidePartPath);
                    leaf = customTextRectangleBottomGuideEntries[0];
                    attribute = "fmla";
                }
                else
                {
                    leaf = customTextRectangleBottom;
                    attribute = "b";
                }
                break;
            case "presetGeometryAdjustment":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose a preset-geometry adjustment.", operation.SlidePartPath);
                var presetGeometry = DirectChildRange(xml, properties, "spPr", "prstGeom", operation);
                var presetAdjustmentList = DirectChildRange(xml, presetGeometry, "prstGeom", "avLst", operation);
                var presetAdjustmentGuides = DirectChildRanges(xml, presetAdjustmentList)
                    .Where(entry => entry.LocalName == "gd")
                    .ToArray();
                if (operation.NativeLeafIndex >= (uint)presetAdjustmentGuides.Length)
                    throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} preset-geometry adjustment index is out of range.", operation.SlidePartPath);
                leaf = presetAdjustmentGuides[(int)operation.NativeLeafIndex];
                attribute = "fmla";
                break;
            case "fillScheme":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose fillScheme.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, "spPr", "solidFill", operation), "solidFill", "schemeClr", operation);
                attribute = "val";
                break;
            case "lineRgb":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineRgb.", operation.SlidePartPath);
                var outline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                if (owner == "cxnSp" && DirectChildRanges(xml, outline).Count(entry => entry.LocalName == "solidFill") == 0)
                    return CompileNativeConnectorStyleColorXmlPatch(xml, elementRange, proof);
                leaf = DirectChildRange(xml, DirectChildRange(xml, outline, "ln", "solidFill", operation), "solidFill", "srgbClr", operation);
                attribute = "val";
                break;
            case "lineScheme":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineScheme.", operation.SlidePartPath);
                var schemeOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                if (owner == "cxnSp" && DirectChildRanges(xml, schemeOutline).Count(entry => entry.LocalName == "solidFill") == 0)
                    return CompileNativeConnectorStyleColorXmlPatch(xml, elementRange, proof);
                leaf = DirectChildRange(xml, DirectChildRange(xml, schemeOutline, "ln", "solidFill", operation), "solidFill", "schemeClr", operation);
                attribute = "val";
                break;
            case "lineOpacityThousandthPercent":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose line opacity.", operation.SlidePartPath);
                var lineOpacityOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                var lineOpacitySolidFill = DirectChildRange(xml, lineOpacityOutline, "ln", "solidFill", operation);
                var lineOpacityColors = DirectChildRanges(xml, lineOpacitySolidFill)
                    .Where(entry => entry.LocalName is "srgbClr" or "schemeClr")
                    .ToArray();
                if (lineOpacityColors.Length != 1)
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} line color is missing or ambiguous.", operation.SlidePartPath);
                var lineOpacityColor = lineOpacityColors[0];
                leaf = DirectChildRange(xml, lineOpacityColor, lineOpacityColor.LocalName, "alpha", operation);
                attribute = "val";
                break;
            case "lineStyle":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineStyle.", operation.SlidePartPath);
                var styleOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                leaf = DirectChildRange(xml, styleOutline, "ln", "prstDash", operation);
                attribute = "val";
                break;
            case "lineCap":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineCap.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "ln", operation);
                attribute = "cap";
                break;
            case "lineJoin":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineJoin.", operation.SlidePartPath);
                var joinOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                return CompileNativeJoinXmlPatch(xml, joinOutline, proof);
            case "lineStartArrow":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineStartArrow.", operation.SlidePartPath);
                var startArrowOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                return CompileNativeArrowXmlPatch(xml, startArrowOutline, proof, "headEnd");
            case "lineEndArrow":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineEndArrow.", operation.SlidePartPath);
                var endArrowOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                return CompileNativeArrowXmlPatch(xml, endArrowOutline, proof, "tailEnd");
            case "lineStartArrowWidth":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineStartArrowWidth.", operation.SlidePartPath);
                return CompileNativeArrowSizeXmlPatch(xml, DirectChildRange(xml, properties, "spPr", "ln", operation), proof, "headEnd", "w");
            case "lineStartArrowLength":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineStartArrowLength.", operation.SlidePartPath);
                return CompileNativeArrowSizeXmlPatch(xml, DirectChildRange(xml, properties, "spPr", "ln", operation), proof, "headEnd", "len");
            case "lineEndArrowWidth":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineEndArrowWidth.", operation.SlidePartPath);
                return CompileNativeArrowSizeXmlPatch(xml, DirectChildRange(xml, properties, "spPr", "ln", operation), proof, "tailEnd", "w");
            case "lineEndArrowLength":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineEndArrowLength.", operation.SlidePartPath);
                return CompileNativeArrowSizeXmlPatch(xml, DirectChildRange(xml, properties, "spPr", "ln", operation), proof, "tailEnd", "len");
            case "lineWidthEmu":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineWidthEmu.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "ln", operation);
                attribute = "w";
                break;
            case "leftEmu":
            case "topEmu":
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, propertiesName, "xfrm", operation), "xfrm", "off", operation);
                attribute = LeafKind(operation) == "leftEmu" ? "x" : "y";
                break;
            case "widthEmu":
            case "heightEmu":
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, propertiesName, "xfrm", operation), "xfrm", "ext", operation);
                attribute = LeafKind(operation) == "widthEmu" ? "cx" : "cy";
                break;
            case "childLeftEmu":
            case "childTopEmu":
                if (owner != "grpSp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, propertiesName, "xfrm", operation), "xfrm", "chOff", operation);
                attribute = LeafKind(operation) == "childLeftEmu" ? "x" : "y";
                break;
            case "childWidthEmu":
            case "childHeightEmu":
                if (owner != "grpSp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, propertiesName, "xfrm", operation), "xfrm", "chExt", operation);
                attribute = LeafKind(operation) == "childWidthEmu" ? "cx" : "cy";
                break;
            case "rotationDegrees":
                if (owner is not ("sp" or "pic" or "grpSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose rotationDegrees.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, propertiesName, "xfrm", operation);
                attribute = "rot";
                break;
            case "flipHorizontal":
            case "flipVertical":
                if (owner is not ("sp" or "pic" or "grpSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, propertiesName, "xfrm", operation);
                attribute = LeafKind(operation) == "flipHorizontal" ? "flipH" : "flipV";
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
        var expectedScalar = LeafKind(operation) switch
        {
            "lineStyle" => PptxLineStyleCodec.TryPresetDashToken(operation.ExpectedValue, out var expectedStyleToken) ? expectedStyleToken : string.Empty,
            "lineCap" => PptxLineStyleCodec.TryCapToken(operation.ExpectedValue, out var expectedCapToken) ? expectedCapToken : string.Empty,
            "imageMaskPreset" => PptxCustomGeometryCodec.TryPreset(operation.ExpectedValue, out _) ? operation.ExpectedValue : string.Empty,
            "imageMaskAdjustment" => $"val {operation.ExpectedValue}",
            "customGeometryAdjustment" => $"val {operation.ExpectedValue}",
            "customGeometryAdjustmentFormula" => operation.ExpectedValue,
            "customGeometryGuide" => $"val {operation.ExpectedValue}",
            "customGeometryGuideFormula" => operation.ExpectedValue,
            "customGeometryPathFill" => operation.ExpectedValue == "1" ? "norm" : "none",
            "presetGeometryAdjustment" => $"val {operation.ExpectedValue}",
            "shape3dContourRgb" => PptxColor.Normalize(operation.ExpectedValue),
            "shape3dContourColorScheme" => PptxColor.NormalizeScheme(operation.ExpectedValue),
            "shape3dExtrusionRgb" => PptxColor.Normalize(operation.ExpectedValue),
            "shape3dExtrusionColorScheme" => PptxColor.NormalizeScheme(operation.ExpectedValue),
            _ => operation.ExpectedValue,
        };
        var replacement = LeafKind(operation) switch
        {
            "lineStyle" => PptxLineStyleCodec.TryPresetDashToken(operation.Value, out var requestedStyleToken) ? requestedStyleToken : string.Empty,
            "lineCap" => PptxLineStyleCodec.TryCapToken(operation.Value, out var requestedCapToken) ? requestedCapToken : string.Empty,
            "imageMaskPreset" => PptxCustomGeometryCodec.TryPreset(operation.Value, out _) ? operation.Value : string.Empty,
            "imageMaskAdjustment" => $"val {operation.Value}",
            "customGeometryAdjustment" => $"val {operation.Value}",
            "customGeometryAdjustmentFormula" => operation.Value,
            "customGeometryGuide" => $"val {operation.Value}",
            "customGeometryGuideFormula" => operation.Value,
            "customGeometryPathFill" => operation.Value == "1" ? "norm" : "none",
            "presetGeometryAdjustment" => $"val {operation.Value}",
            "shape3dContourRgb" => PptxColor.Normalize(operation.Value),
            "shape3dContourColorScheme" => PptxColor.NormalizeScheme(operation.Value),
            "shape3dExtrusionRgb" => PptxColor.Normalize(operation.Value),
            "shape3dExtrusionColorScheme" => PptxColor.NormalizeScheme(operation.Value),
            _ => operation.Value,
        };
        var privateTextRectangleEdge =
            (LeafKind(operation) is "customGeometryTextRectangleLeftEmu" or "customGeometryTextRectangleTopEmu" or "customGeometryTextRectangleRightEmu" or "customGeometryTextRectangleBottomEmu") && leaf.LocalName == "gd";
        var privateTextRectangleAxis = string.Empty;
        var privateTextRectangleExtent = string.Empty;
        if (privateTextRectangleEdge && !TryTextRectanglePrivateGuideFormula(
                valueGroup.Value, operation.ExpectedValue,
                out privateTextRectangleAxis, out privateTextRectangleExtent))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} private text-rectangle guide formula is not canonical or does not match the expected coordinate.", operation.SlidePartPath);
        if (privateTextRectangleEdge)
        {
            expectedScalar = valueGroup.Value;
            replacement = $"*/ {operation.Value} {privateTextRectangleAxis} {privateTextRectangleExtent}";
        }
        var matches = LeafKind(operation) is "flipHorizontal" or "flipVertical" or "tableBandedRows" or "tableBandedColumns" or "tableFirstColumnEmphasis" or "tableLastColumnEmphasis" or "tableLastRow" or "customGeometryPathStroke" or "customGeometryPathExtrusionAllowed"
            ? TryCanonicalBoolean(valueGroup.Value, out var actualBoolean) && actualBoolean == operation.ExpectedValue
            : LeafKind(operation) is "shape3dContourRgb" or "shape3dExtrusionRgb"
                ? IsRgbToken(valueGroup.Value) && PptxColor.Normalize(valueGroup.Value) == expectedScalar
                : LeafKind(operation) is "shape3dContourColorScheme" or "shape3dExtrusionColorScheme"
                    ? PptxColor.TrySchemeToken(valueGroup.Value, out var actualScheme) && PptxColor.NormalizeScheme(actualScheme) == expectedScalar
                : valueGroup.Value == expectedScalar;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw scalar does not match the expected value.", operation.SlidePartPath);
        var start = leaf.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTableBooleanXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        if (elementRange.LocalName != "graphicFrame")
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {LeafKind(operation)} target has the wrong native element type.", operation.SlidePartPath);
        var graphic = DirectChildRange(xml, elementRange, "graphicFrame", "graphic", operation);
        var graphicData = DirectChildRange(xml, graphic, "graphic", "graphicData", operation);
        var table = DirectChildRange(xml, graphicData, "graphicData", "tbl", operation);
        var properties = DirectChildRange(xml, table, "tbl", "tblPr", operation);
        var fragment = xml[properties.Start..properties.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == properties.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} table property tag was not found.", operation.SlidePartPath);
        var attributeName = LeafKind(operation) switch
        {
            "tableHeaderRows" => "firstRow",
            "tableBandedRows" => "bandRow",
            "tableFirstColumnEmphasis" => "firstCol",
            "tableLastColumnEmphasis" => "lastCol",
            "tableLastRow" => "lastRow",
            _ => "bandCol",
        };
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attributeName)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} table {attributeName} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (!TryCanonicalBoolean(valueGroup.Value, out var actual) || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw table {attributeName} does not match the expected value.", operation.SlidePartPath);
        var start = properties.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileGlowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var owner = kind.StartsWith("shapeGlow", StringComparison.Ordinal) ? "sp" : "pic";
        if (elementRange.LocalName != owner)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
        var properties = DirectChildRange(xml, elementRange, owner, "spPr", operation);
        var effectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
        var glow = DirectChildRange(xml, effectList, "effectLst", "glow", operation);
        XmlRange target;
        string attribute;
        string expectedScalar;
        string replacement;
        if (kind is "shapeGlowRadiusEmu" or "imageGlowRadiusEmu")
        {
            target = glow;
            attribute = "rad";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }
        else if (kind is "shapeGlowColorRgb" or "shapeGlowColorScheme" or "imageGlowColorRgb" or "imageGlowColorScheme")
        {
            var colorName = kind is "shapeGlowColorRgb" or "imageGlowColorRgb" ? "srgbClr" : "schemeClr";
            target = DirectChildRange(xml, glow, "glow", colorName, operation);
            attribute = "val";
            expectedScalar = kind is "shapeGlowColorRgb" or "imageGlowColorRgb"
                ? PptxColor.Normalize(operation.ExpectedValue)
                : PptxColor.NormalizeScheme(operation.ExpectedValue);
            replacement = kind is "shapeGlowColorRgb" or "imageGlowColorRgb"
                ? PptxColor.Normalize(operation.Value)
                : PptxColor.NormalizeScheme(operation.Value);
        }
        else
        {
            var colors = DirectChildRanges(xml, glow)
                .Where(entry => entry.LocalName is "srgbClr" or "schemeClr")
                .ToArray();
            if (colors.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} glow color is missing or ambiguous.", operation.SlidePartPath);
            target = DirectChildRange(xml, colors[0], colors[0].LocalName, "alpha", operation);
            attribute = "val";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }

        var fragment = xml[target.Start..target.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == target.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} glow scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} glow scalar attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var isRgb = kind is "shapeGlowColorRgb" or "imageGlowColorRgb";
        var isScheme = kind is "shapeGlowColorScheme" or "imageGlowColorScheme";
        var matches = isRgb
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expectedScalar
            : isScheme
                ? PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expectedScalar
                : actual == expectedScalar;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw glow value does not match the expected value.", operation.SlidePartPath);
        var start = target.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileInnerShadowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var owner = kind.StartsWith("shapeInnerShadow", StringComparison.Ordinal) ? "sp" : "pic";
        if (elementRange.LocalName != owner)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
        var properties = DirectChildRange(xml, elementRange, owner, "spPr", operation);
        var effectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
        var innerShadow = DirectChildRange(xml, effectList, "effectLst", "innerShdw", operation);
        XmlRange target;
        string attribute;
        string expectedScalar;
        string replacement;
        if (kind is "shapeInnerShadowBlurRadiusEmu" or "imageInnerShadowBlurRadiusEmu" or
            "shapeInnerShadowDistanceEmu" or "imageInnerShadowDistanceEmu" or
            "shapeInnerShadowDirectionDegrees" or "imageInnerShadowDirectionDegrees")
        {
            target = innerShadow;
            attribute = kind switch
            {
                "shapeInnerShadowBlurRadiusEmu" or "imageInnerShadowBlurRadiusEmu" => "blurRad",
                "shapeInnerShadowDistanceEmu" or "imageInnerShadowDistanceEmu" => "dist",
                _ => "dir",
            };
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }
        else if (kind is "shapeInnerShadowColorRgb" or "shapeInnerShadowColorScheme" or
                 "imageInnerShadowColorRgb" or "imageInnerShadowColorScheme")
        {
            var colorName = kind is "shapeInnerShadowColorRgb" or "imageInnerShadowColorRgb" ? "srgbClr" : "schemeClr";
            target = DirectChildRange(xml, innerShadow, "innerShdw", colorName, operation);
            attribute = "val";
            expectedScalar = kind is "shapeInnerShadowColorRgb" or "imageInnerShadowColorRgb"
                ? PptxColor.Normalize(operation.ExpectedValue)
                : PptxColor.NormalizeScheme(operation.ExpectedValue);
            replacement = kind is "shapeInnerShadowColorRgb" or "imageInnerShadowColorRgb"
                ? PptxColor.Normalize(operation.Value)
                : PptxColor.NormalizeScheme(operation.Value);
        }
        else
        {
            var colors = DirectChildRanges(xml, innerShadow)
                .Where(entry => entry.LocalName is "srgbClr" or "schemeClr")
                .ToArray();
            if (colors.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} inner-shadow color is missing or ambiguous.", operation.SlidePartPath);
            target = DirectChildRange(xml, colors[0], colors[0].LocalName, "alpha", operation);
            attribute = "val";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }

        var fragment = xml[target.Start..target.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == target.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} inner-shadow scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} inner-shadow scalar attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var isRgb = kind is "shapeInnerShadowColorRgb" or "imageInnerShadowColorRgb";
        var isScheme = kind is "shapeInnerShadowColorScheme" or "imageInnerShadowColorScheme";
        var matches = isRgb
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expectedScalar
            : isScheme
                ? PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expectedScalar
                : actual == expectedScalar;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw inner-shadow value does not match the expected value.", operation.SlidePartPath);
        var start = target.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileReflectionXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var owner = kind.StartsWith("shapeReflection", StringComparison.Ordinal) ? "sp" : "pic";
        if (elementRange.LocalName != owner)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
        var properties = DirectChildRange(xml, elementRange, owner, "spPr", operation);
        var effectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
        var reflection = DirectChildRange(xml, effectList, "effectLst", "reflection", operation);
        var attribute = kind switch
        {
            "shapeReflectionBlurRadiusEmu" or "imageReflectionBlurRadiusEmu" => "blurRad",
            "shapeReflectionStartOpacityThousandthPercent" or "imageReflectionStartOpacityThousandthPercent" => "stA",
            "shapeReflectionEndOpacityThousandthPercent" or "imageReflectionEndOpacityThousandthPercent" => "endA",
            "shapeReflectionDistanceEmu" or "imageReflectionDistanceEmu" => "dist",
            "shapeReflectionDirectionDegrees" or "imageReflectionDirectionDegrees" => "dir",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported reflection leaf.", operation.SlidePartPath),
        };
        var fragment = xml[reflection.Start..reflection.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "reflection")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} reflection scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} reflection scalar attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        if (actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw reflection value does not match the expected value.", operation.SlidePartPath);
        var start = reflection.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileSoftEdgeXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var owner = kind == "shapeSoftEdgeRadiusEmu" ? "sp" : "pic";
        if (elementRange.LocalName != owner)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target has the wrong native element type.", operation.SlidePartPath);
        var properties = DirectChildRange(xml, elementRange, owner, "spPr", operation);
        var effectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
        var softEdge = DirectChildRange(xml, effectList, "effectLst", "softEdge", operation);
        var fragment = xml[softEdge.Start..softEdge.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "softEdge")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} soft-edge scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "rad")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} soft-edge scalar attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (System.Net.WebUtility.HtmlDecode(valueGroup.Value) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw soft-edge value does not match the expected value.", operation.SlidePartPath);
        var start = softEdge.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextParagraphAlignmentXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var paragraphs = ShapeElementRanges(elementXml, "txBody")
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph-alignment index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var pPrXml = elementXml[pPr.Start..pPr.End];
        var startTag = XmlTokenPattern().Matches(pPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "pPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "algn")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph alignment attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParagraphAlignmentName(System.Net.WebUtility.HtmlDecode(valueGroup.Value));
        if (actual.Length == 0 || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph alignment does not match the expected value.", operation.SlidePartPath);
        var replacement = ParagraphAlignmentToken(operation.Value);
        var start = elementRange.Start + pPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextParagraphSpacingXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var paragraphs = ShapeElementRanges(elementXml, "txBody")
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph spacing index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var kind = LeafKind(operation);
        var spacingName = kind switch
        {
            "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" => "lnSpc",
            "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" => "spcBef",
            "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier" => "spcAft",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph spacing leaf.", operation.SlidePartPath),
        };
        var spacing = DirectChildRange(elementXml, pPr, "pPr", spacingName, operation);
        var children = DirectChildRanges(elementXml, spacing);
        var expectedName = kind.EndsWith("Points", StringComparison.Ordinal) ? "spcPts" : "spcPct";
        if (children.Count != 1 || children[0].LocalName != expectedName)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph spacing has the wrong native unit or ambiguous child.", operation.SlidePartPath);
        var child = children[0];
        var childXml = elementXml[child.Start..child.End];
        var startTag = XmlTokenPattern().Matches(childXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == expectedName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph spacing child tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph spacing value attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseParagraphSpacingToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), kind, operation);
        var expected = ParseParagraphSpacingToken(operation.ExpectedValue, kind, operation);
        if (actual != expected)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph spacing does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseParagraphSpacingToken(operation.Value, kind, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var start = elementRange.Start + child.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextParagraphLayoutXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var paragraphs = ShapeElementRanges(elementXml, "txBody")
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph layout index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var kind = LeafKind(operation);
        var attributeName = kind switch
        {
            "paragraphMarginLeftEmu" => "marL",
            "paragraphIndentEmu" => "indent",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph layout leaf.", operation.SlidePartPath),
        };
        var pPrXml = elementXml[pPr.Start..pPr.End];
        var startTag = XmlTokenPattern().Matches(pPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "pPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attributeName)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph {attributeName} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseParagraphLayoutToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), kind, operation);
        var expected = ParseParagraphLayoutToken(operation.ExpectedValue, kind, operation);
        if (actual != expected)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph layout does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseParagraphLayoutToken(operation.Value, kind, operation)
            .ToString(System.Globalization.CultureInfo.InvariantCulture);
        var start = elementRange.Start + pPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextParagraphBulletCharacterXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var paragraphs = ShapeElementRanges(elementXml, "txBody")
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph bullet index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var bullets = DirectChildRanges(elementXml, pPr)
            .Where(range => range.LocalName == "buChar")
            .ToArray();
        if (bullets.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct buChar child under pPr.", operation.SlidePartPath);
        var bulletXml = elementXml[bullets[0].Start..bullets[0].End];
        var startTag = XmlTokenPattern().Matches(bulletXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "buChar")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} character bullet tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "char")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} character bullet attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        if (!TryBulletCharacter(actual) || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw character bullet does not match the expected value.", operation.SlidePartPath);
        if (!TryBulletCharacter(operation.Value))
            throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} paragraphBulletCharacter must contain one Unicode scalar value.", operation.SlidePartPath);
        var replacement = EscapeXmlAttribute(operation.Value);
        var start = elementRange.Start + bullets[0].Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextParagraphAutoNumberXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var paragraphs = ShapeElementRanges(elementXml, "txBody")
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph auto-number index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var markers = DirectChildRanges(elementXml, pPr)
            .Where(range => range.LocalName == "buAutoNum")
            .ToArray();
        if (markers.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct buAutoNum child under pPr.", operation.SlidePartPath);
        var marker = markers[0];
        var markerXml = elementXml[marker.Start..marker.End];
        var startTag = XmlTokenPattern().Matches(markerXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "buAutoNum")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} auto-number tag was not found.", operation.SlidePartPath);
        if (!startTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) || DirectChildRanges(markerXml, new XmlRange(0, markerXml.Length, "buAutoNum")).Count != 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires a bare self-closing buAutoNum element.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>().ToArray();
        var invalidAttribute = attributes.Any(attribute => LocalAttributeName(attribute.Groups["name"].Value) is not ("type" or "startAt"));
        var typeAttributes = attributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "type").ToArray();
        var startAttributes = attributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "startAt").ToArray();
        if (invalidAttribute || typeAttributes.Length != 1 || startAttributes.Length > 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} auto-number attributes are missing or ambiguous.", operation.SlidePartPath);
        var typeValue = System.Net.WebUtility.HtmlDecode(typeAttributes[0].Groups["value"].Value);
        if (!PptxBulletCodec.IsAutoNumberScheme(typeValue))
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no supported auto-number scheme.", operation.SlidePartPath);
        var kind = LeafKind(operation);
        Match selected;
        string expected;
        string replacement;
        if (kind == "paragraphBulletAutoNumberScheme")
        {
            selected = typeAttributes[0];
            expected = operation.ExpectedValue;
            replacement = operation.Value;
            if (!PptxBulletCodec.IsAutoNumberScheme(expected) || !PptxBulletCodec.IsAutoNumberScheme(replacement) || !typeValue.Equals(expected, StringComparison.Ordinal))
                throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw auto-number scheme does not match the expected value.", operation.SlidePartPath);
        }
        else
        {
            selected = startAttributes.FirstOrDefault() ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no explicit auto-number start.", operation.SlidePartPath);
            expected = ParseParagraphAutoNumberStartAtToken(operation.ExpectedValue, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
            replacement = ParseParagraphAutoNumberStartAtToken(operation.Value, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
            var actual = ParseParagraphAutoNumberStartAtToken(System.Net.WebUtility.HtmlDecode(selected.Groups["value"].Value), operation);
            if (actual.ToString(System.Globalization.CultureInfo.InvariantCulture) != expected)
                throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw auto-number start does not match the expected value.", operation.SlidePartPath);
        }
        var valueGroup = selected.Groups["value"];
        var start = elementRange.Start + marker.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, EscapeXmlAttribute(replacement), proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextParagraphBulletStyleXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var paragraphs = ShapeElementRanges(elementXml, "txBody")
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph bullet-style index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var kind = LeafKind(operation);
        var markerName = kind switch
        {
            "paragraphBulletFontFamily" => "buFont",
            "paragraphBulletColorRgb" or "paragraphBulletColorScheme" => "buClr",
            "paragraphBulletSizePoints" => "buSzPts",
            "paragraphBulletSizePercent" => "buSzPct",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph bullet-style leaf.", operation.SlidePartPath),
        };
        var markers = DirectChildRanges(elementXml, pPr)
            .Where(range => range.LocalName == markerName)
            .ToArray();
        if (markers.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct {markerName} child under pPr.", operation.SlidePartPath);
        var marker = markers[0];
        var markerXml = elementXml[marker.Start..marker.End];
        var markerTag = XmlTokenPattern().Matches(markerXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == markerName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} bullet-style tag was not found.", operation.SlidePartPath);
        var markerAttributes = XmlAttributePattern().Matches(markerTag.Value).Cast<Match>().ToArray();
        Match selected;
        string expected;
        string replacement;
        int valueStart;
        if (kind == "paragraphBulletFontFamily" || kind is "paragraphBulletSizePoints" or "paragraphBulletSizePercent")
        {
            if (!markerTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) || DirectChildRanges(markerXml, new XmlRange(0, markerXml.Length, markerName)).Count != 0)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires a bare self-closing {markerName} element.", operation.SlidePartPath);
            var attributeName = kind == "paragraphBulletFontFamily" ? "typeface" : "val";
            var attributes = markerAttributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == attributeName).ToArray();
            if (markerAttributes.Any(attribute => LocalAttributeName(attribute.Groups["name"].Value) != attributeName) || attributes.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {markerName} attributes are missing or ambiguous.", operation.SlidePartPath);
            selected = attributes[0];
            expected = kind == "paragraphBulletFontFamily"
                ? operation.ExpectedValue
                : ParseParagraphBulletSizeToken(operation.ExpectedValue, kind, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
            replacement = kind == "paragraphBulletFontFamily"
                ? operation.Value
                : ParseParagraphBulletSizeToken(operation.Value, kind, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
            var actual = System.Net.WebUtility.HtmlDecode(selected.Groups["value"].Value);
            if (kind == "paragraphBulletFontFamily")
            {
                if (!ValidFontFamilyToken(actual) || !actual.Equals(expected, StringComparison.Ordinal))
                    throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw bullet font does not match the expected value.", operation.SlidePartPath);
            }
            else if (ParseParagraphBulletSizeToken(actual, kind, operation).ToString(System.Globalization.CultureInfo.InvariantCulture) != expected)
            {
                throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw bullet size does not match the expected value.", operation.SlidePartPath);
            }
            valueStart = marker.Start + markerTag.Index;
        }
        else
        {
            if (!markerTag.Value.EndsWith(">", StringComparison.Ordinal) || markerTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) || markerAttributes.Length != 0)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires an attribute-free {markerName} wrapper.", operation.SlidePartPath);
            var children = DirectChildRanges(elementXml, marker);
            if (children.Count != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct color child under buClr.", operation.SlidePartPath);
            var child = children[0];
            var childName = kind == "paragraphBulletColorRgb" ? "srgbClr" : "schemeClr";
            if (child.LocalName != childName)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} bullet color uses the wrong color model.", operation.SlidePartPath);
            var childXml = elementXml[child.Start..child.End];
            var childTag = XmlTokenPattern().Matches(childXml).Cast<Match>()
                .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == childName)
                ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} bullet color child tag was not found.", operation.SlidePartPath);
            if (!childTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires a bare self-closing bullet color child.", operation.SlidePartPath);
            var valueAttributes = XmlAttributePattern().Matches(childTag.Value).Cast<Match>()
                .Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "val").ToArray();
            var allAttributes = XmlAttributePattern().Matches(childTag.Value).Cast<Match>().ToArray();
            if (allAttributes.Length != 1 || valueAttributes.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} bullet color value is missing or ambiguous.", operation.SlidePartPath);
            selected = valueAttributes[0];
            expected = kind == "paragraphBulletColorRgb" ? PptxColor.Normalize(operation.ExpectedValue) : PptxColor.NormalizeScheme(operation.ExpectedValue);
            replacement = kind == "paragraphBulletColorRgb" ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
            var actual = System.Net.WebUtility.HtmlDecode(selected.Groups["value"].Value);
            if (kind == "paragraphBulletColorRgb")
            {
                if (PptxColor.Normalize(actual) != expected)
                    throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw bullet color does not match the expected value.", operation.SlidePartPath);
            }
            else if (!PptxColor.TrySchemeToken(actual, out var actualScheme) || actualScheme != expected)
            {
                throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw bullet color does not match the expected value.", operation.SlidePartPath);
            }
            valueStart = child.Start + childTag.Index;
        }
        var valueGroup = selected.Groups["value"];
        var start = elementRange.Start + valueStart + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, EscapeXmlAttribute(replacement), proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextParagraphLevelXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var paragraphs = ShapeElementRanges(elementXml, "txBody")
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph level index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var pPrXml = elementXml[pPr.Start..pPr.End];
        var startTag = XmlTokenPattern().Matches(pPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "pPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "lvl")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph level attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseParagraphLevelToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        var expected = ParseParagraphLevelToken(operation.ExpectedValue, operation);
        if (actual != expected)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph level does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseParagraphLevelToken(operation.Value, operation)
            .ToString(System.Globalization.CultureInfo.InvariantCulture);
        var start = elementRange.Start + pPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextVerticalAnchorXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "anchor")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} vertical anchor attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParagraphVerticalAnchorName(System.Net.WebUtility.HtmlDecode(valueGroup.Value));
        if (actual.Length == 0 || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw vertical anchor does not match the expected value.", operation.SlidePartPath);
        var replacement = ParagraphVerticalAnchorToken(operation.Value);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyInsetXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var attributeName = TextBodyInsetAttribute(LeafKind(operation));
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attributeName)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body inset attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (!long.TryParse(System.Net.WebUtility.HtmlDecode(valueGroup.Value), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var actual) || actual < 0 || actual > int.MaxValue || actual.ToString(System.Globalization.CultureInfo.InvariantCulture) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body inset does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseBoundedInsetToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyWrapXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "wrap")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body wrap attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = TextBodyWrapName(System.Net.WebUtility.HtmlDecode(valueGroup.Value));
        if (actual.Length == 0 || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body wrap does not match the expected value.", operation.SlidePartPath);
        var replacement = TextBodyWrapToken(operation.Value);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyColumnCountXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "numCol")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body column-count attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyColumnCountToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.ToString(System.Globalization.CultureInfo.InvariantCulture).Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body column count does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyColumnCountToken(operation.Value, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyColumnGapXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "spcCol")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body column gap attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyColumnGapToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.ToString(System.Globalization.CultureInfo.InvariantCulture).Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body column gap does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyColumnGapToken(operation.Value, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyRotationXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "rot")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body rotation attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyRotationToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.ToString(System.Globalization.CultureInfo.InvariantCulture).Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body rotation does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyRotationToken(operation.Value, operation).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyVerticalOverflowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "vertOverflow")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body vertical overflow attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyVerticalOverflowToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body vertical overflow does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyVerticalOverflowToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyHorizontalOverflowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "horzOverflow")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body horizontal overflow attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyHorizontalOverflowToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body horizontal overflow does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyHorizontalOverflowToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyUprightXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "upright")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body upright attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyUprightToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body upright flag does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyUprightToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyAnchorCenterXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "anchorCtr")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body anchor-center attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyAnchorCenterToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body anchor-center flag does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyAnchorCenterToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyForceAntiAliasXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "forceAA")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body force anti-alias attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyForceAntiAliasToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body force anti-alias flag does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyForceAntiAliasToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodySpaceFirstLastParagraphXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "spcFirstLastPara")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body first-last paragraph spacing attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodySpaceFirstLastParagraphToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body first-last paragraph spacing flag does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodySpaceFirstLastParagraphToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyCompatibleLineSpacingXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "compatLnSpc")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body compatible line spacing attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyCompatibleLineSpacingToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body compatible line spacing flag does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyCompatibleLineSpacingToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyFromWordArtXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "fromWordArt")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body WordArt marker attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyFromWordArtToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body WordArt marker flag does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyFromWordArtToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyWarpPresetXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var choices = DirectChildRanges(elementXml, bodyPr)
            .Where(child => child.LocalName == "prstTxWarp")
            .ToArray();
        if (choices.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct canonical prstTxWarp child.", operation.SlidePartPath);
        var choice = choices[0];
        var choiceXml = elementXml[choice.Start..choice.End];
        var startTag = XmlTokenPattern().Matches(choiceXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "prstTxWarp")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-warp child tag was not found.", operation.SlidePartPath);
        var adjustmentChildren = DirectChildRanges(choiceXml, new XmlRange(0, choiceXml.Length, choice.LocalName));
        if (adjustmentChildren.Count == 0)
        {
            if (!startTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires a bare self-closing prstTxWarp child.", operation.SlidePartPath);
        }
        else
        {
            RequireTextWarpAdjustmentGuides(choiceXml, choice, operation);
        }
        var allAttributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>().ToArray();
        var attributes = allAttributes
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "prst")
            .ToArray();
        if (allAttributes.Length != 1 || attributes.Length != 1 || allAttributes.Any(match => match.Groups["name"].Value.Contains(':')))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp preset attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyWarpPresetToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body warp preset does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyWarpPresetToken(operation.Value, operation);
        var start = elementRange.Start + choice.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyWarpAdjustmentXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var choices = DirectChildRanges(elementXml, bodyPr)
            .Where(child => child.LocalName == "prstTxWarp")
            .ToArray();
        if (choices.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct canonical prstTxWarp child.", operation.SlidePartPath);
        var choice = choices[0];
        var choiceXml = elementXml[choice.Start..choice.End];
        var startTag = XmlTokenPattern().Matches(choiceXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "prstTxWarp")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-warp child tag was not found.", operation.SlidePartPath);
        var allAttributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>().ToArray();
        var presetAttributes = allAttributes
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "prst")
            .ToArray();
        if (allAttributes.Length != 1 || presetAttributes.Length != 1 || allAttributes.Any(match => match.Groups["name"].Value.Contains(':')))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp preset attribute is missing or ambiguous.", operation.SlidePartPath);
        _ = ParseTextBodyWarpPresetToken(System.Net.WebUtility.HtmlDecode(presetAttributes[0].Groups["value"].Value), operation);
        var adjustmentList = DirectChildRanges(choiceXml, new XmlRange(0, choiceXml.Length, choice.LocalName)).Single();
        var guides = RequireTextWarpAdjustmentGuides(choiceXml, choice, operation);
        if (operation.NativeLeafIndex >= (uint)guides.Count)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-warp adjustment index is out of range.", operation.SlidePartPath);
        var guide = guides[(int)operation.NativeLeafIndex];
        var listXml = choiceXml[adjustmentList.Start..adjustmentList.End];
        var guideXml = listXml[guide.Start..guide.End];
        var guideStartTag = XmlTokenPattern().Matches(guideXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "gd")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-warp adjustment guide tag was not found.", operation.SlidePartPath);
        var formulaAttribute = XmlAttributePattern().Matches(guideStartTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "fmla")
            .ToArray();
        if (formulaAttribute.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment formula is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = formulaAttribute[0].Groups["value"];
        var formula = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        if (!PptxBodyPropertiesCodec.TryLiteralTextWarpAdjustment(formula, out var parsedActual))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment formula is not a canonical literal value.", operation.SlidePartPath);
        var actual = parsedActual.ToString(CultureInfo.InvariantCulture);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body warp adjustment does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyWarpAdjustmentToken(operation.Value, operation);
        var tokenStart = formula.IndexOf(actual, StringComparison.Ordinal);
        if (tokenStart < 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment literal token was not found.", operation.SlidePartPath);
        var start = elementRange.Start + choice.Start + adjustmentList.Start + guide.Start + guideStartTag.Index + valueGroup.Index + tokenStart;
        return new PptxXmlPatch(operation, start, start + actual.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyFlatTextZXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var flatTexts = DirectChildRanges(elementXml, bodyPr)
            .Where(child => child.LocalName == "flatTx")
            .ToArray();
        if (flatTexts.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct canonical flatTx child.", operation.SlidePartPath);
        var flatText = flatTexts[0];
        var flatTextXml = elementXml[flatText.Start..flatText.End];
        var startTag = XmlTokenPattern().Matches(flatTextXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "flatTx")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} flatTx child tag was not found.", operation.SlidePartPath);
        if (!startTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) ||
            DirectChildRanges(flatTextXml, new XmlRange(0, flatTextXml.Length, flatText.LocalName)).Count != 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires a bare self-closing flatTx child.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>().ToArray();
        var zAttributes = attributes
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "z")
            .ToArray();
        if (attributes.Length != 1 || zAttributes.Length != 1 || attributes.Any(match => match.Groups["name"].Value.Contains(':')))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} flatTx z attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = zAttributes[0].Groups["value"];
        var actual = ParseTextBodyFlatTextZToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw flatTx z coordinate does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyFlatTextZToken(operation.Value, operation);
        var start = elementRange.Start + flatText.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static IReadOnlyList<XmlRange> RequireTextWarpAdjustmentGuides(
        string choiceXml,
        XmlRange choice,
        PresentationEditOperation operation)
    {
        var children = DirectChildRanges(choiceXml, new XmlRange(0, choiceXml.Length, choice.LocalName));
        if (children.Count != 1 || children[0].LocalName != "avLst")
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct avLst under prstTxWarp.", operation.SlidePartPath);
        var list = children[0];
        var listXml = choiceXml[list.Start..list.End];
        var listStartTag = XmlTokenPattern().Matches(listXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "avLst")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-warp adjustment list tag was not found.", operation.SlidePartPath);
        if (XmlAttributePattern().Matches(listStartTag.Value).Count != 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment list has unsupported attributes.", operation.SlidePartPath);
        var guides = DirectChildRanges(listXml, new XmlRange(0, listXml.Length, list.LocalName));
        if (guides.Count == 0 || guides.Count > 256 || guides.Any(guide => guide.LocalName != "gd"))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment list is empty or has unsupported children.", operation.SlidePartPath);
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var guide in guides)
        {
            var guideXml = listXml[guide.Start..guide.End];
            var guideStartTag = XmlTokenPattern().Matches(guideXml).Cast<Match>()
                .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "gd")
                ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-warp adjustment guide tag was not found.", operation.SlidePartPath);
            if (!guideStartTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) ||
                DirectChildRanges(guideXml, new XmlRange(0, guideXml.Length, guide.LocalName)).Count != 0)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment guides must be bare self-closing elements.", operation.SlidePartPath);
            var attributes = XmlAttributePattern().Matches(guideStartTag.Value).Cast<Match>().ToArray();
            var nameAttributes = attributes.Where(match => LocalAttributeName(match.Groups["name"].Value) == "name").ToArray();
            var formulaAttributes = attributes.Where(match => LocalAttributeName(match.Groups["name"].Value) == "fmla").ToArray();
            if (attributes.Length != 2 || nameAttributes.Length != 1 || formulaAttributes.Length != 1 ||
                attributes.Any(match => match.Groups["name"].Value.Contains(':')))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment guide attributes are missing or ambiguous.", operation.SlidePartPath);
            var name = System.Net.WebUtility.HtmlDecode(nameAttributes[0].Groups["value"].Value);
            var formula = System.Net.WebUtility.HtmlDecode(formulaAttributes[0].Groups["value"].Value);
            if (name.Length == 0 || name.Length > 256 || name.Any(char.IsControl) || !names.Add(name) ||
                !PptxBodyPropertiesCodec.TryLiteralTextWarpAdjustment(formula, out _))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-warp adjustment guide is not a canonical literal value.", operation.SlidePartPath);
        }
        return guides;
    }

    private static PptxXmlPatch CompileTextBodyAutoFitXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var choices = DirectChildRanges(elementXml, bodyPr)
            .Where(child => child.LocalName is "noAutofit" or "normAutofit" or "spAutoFit")
            .ToArray();
        if (choices.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct canonical AutoFit child.", operation.SlidePartPath);
        var choice = choices[0];
        var choiceXml = elementXml[choice.Start..choice.End];
        var startTag = XmlTokenPattern().Matches(choiceXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == choice.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} AutoFit child tag was not found.", operation.SlidePartPath);
        if (!startTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) || DirectChildRanges(choiceXml, new XmlRange(0, choiceXml.Length, choice.LocalName)).Count != 0 ||
            XmlAttributePattern().Matches(startTag.Value).Count != 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires a bare self-closing AutoFit child.", operation.SlidePartPath);
        var actual = TextBodyAutoFitName(choice.LocalName);
        if (actual.Length == 0 || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body AutoFit mode does not match the expected value.", operation.SlidePartPath);
        var replacement = TextBodyAutoFitElementName(operation.Value);
        var localMarker = $":{choice.LocalName}";
        var localMarkerIndex = startTag.Value.IndexOf(localMarker, StringComparison.Ordinal);
        if (localMarkerIndex < 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} AutoFit child has no bounded namespace-qualified local name.", operation.SlidePartPath);
        var start = elementRange.Start + choice.Start + startTag.Index + localMarkerIndex + 1;
        return new PptxXmlPatch(operation, start, start + choice.LocalName.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyNormalAutoFitXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var choices = DirectChildRanges(elementXml, bodyPr)
            .Where(child => child.LocalName == "normAutofit")
            .ToArray();
        if (choices.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct canonical normAutofit child.", operation.SlidePartPath);
        var choice = choices[0];
        var choiceXml = elementXml[choice.Start..choice.End];
        var startTag = XmlTokenPattern().Matches(choiceXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "normAutofit")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} normAutofit child tag was not found.", operation.SlidePartPath);
        if (!startTag.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) ||
            DirectChildRanges(choiceXml, new XmlRange(0, choiceXml.Length, choice.LocalName)).Count != 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires a bare self-closing normAutofit child.", operation.SlidePartPath);
        var attributeName = LeafKind(operation) == "textBodyNormalAutoFitFontScale" ? "fontScale" : "lnSpcReduction";
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>().ToArray();
        if (attributes.Any(attribute => attribute.Groups["name"].Value.Contains(':') ||
                                        LocalAttributeName(attribute.Groups["name"].Value) is not ("fontScale" or "lnSpcReduction")))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} normAutofit child has unsupported attributes.", operation.SlidePartPath);
        var selected = attributes
            .Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == attributeName)
            .ToArray();
        if (selected.Length != 1)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} normAutofit {attributeName} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = selected[0].Groups["value"];
        var actual = ParseTextBodyNormalAutoFitToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), LeafKind(operation), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw normal AutoFit percentage does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyNormalAutoFitToken(operation.Value, LeafKind(operation), operation);
        var start = elementRange.Start + choice.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyColumnDirectionXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "rtlCol")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body column direction attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = ParseTextBodyColumnDirectionToken(System.Net.WebUtility.HtmlDecode(valueGroup.Value), operation);
        if (!actual.Equals(operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body column direction does not match the expected value.", operation.SlidePartPath);
        var replacement = ParseTextBodyColumnDirectionToken(operation.Value, operation);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextBodyVerticalTextXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var bodyPr = DirectChildRange(elementXml, txBody, "txBody", "bodyPr", operation);
        var bodyPrXml = elementXml[bodyPr.Start..bodyPr.End];
        var startTag = XmlTokenPattern().Matches(bodyPrXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "bodyPr")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} body properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "vert")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text-body vertical text attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = TextBodyVerticalTextName(System.Net.WebUtility.HtmlDecode(valueGroup.Value));
        if (actual.Length == 0 || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text-body vertical text mode does not match the expected value.", operation.SlidePartPath);
        var replacement = TextBodyVerticalTextToken(operation.Value);
        var start = elementRange.Start + bodyPr.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static string TextBodyInsetAttribute(string leafKind) => leafKind switch
    {
        "textBodyInsetLeftEmu" => "lIns",
        "textBodyInsetTopEmu" => "tIns",
        "textBodyInsetRightEmu" => "rIns",
        "textBodyInsetBottomEmu" => "bIns",
        _ => throw new CodecException("invalid_presentation_edit_target", $"Unsupported Presentation text-body inset leaf {leafKind}."),
    };

    private static string TextBodyWrapName(A.TextWrappingValues? value) => value is { } current && current == A.TextWrappingValues.Square
        ? "square"
        : value is { } none && none == A.TextWrappingValues.None
            ? "none"
            : string.Empty;

    private static string TextBodyWrapName(string value) => value switch
    {
        "square" => "square",
        "none" => "none",
        _ => string.Empty,
    };

    private static string TextBodyWrapToken(string value) => value switch
    {
        "square" => "square",
        "none" => "none",
        _ => throw new CodecException("invalid_presentation_edit_target", $"Unsupported Presentation text-body wrap mode {value}."),
    };

    private static string TextBodyVerticalTextName(A.TextVerticalValues? value) => value is { } current && current == A.TextVerticalValues.Horizontal
        ? "horizontal"
        : value is { } vertical && vertical == A.TextVerticalValues.Vertical
            ? "vertical"
            : value is { } vertical270 && vertical270 == A.TextVerticalValues.Vertical270
                ? "vertical270"
                : string.Empty;

    private static string TextBodyVerticalTextName(string value) => value switch
    {
        "horz" => "horizontal",
        "vert" => "vertical",
        "vert270" => "vertical270",
        _ => string.Empty,
    };

    private static string TextBodyVerticalTextToken(string value) => value switch
    {
        "horizontal" => "horz",
        "vertical" => "vert",
        "vertical270" => "vert270",
        _ => throw new CodecException("invalid_presentation_edit_target", $"Unsupported Presentation text-body vertical text mode {value}."),
    };

    private static string TextBodyAutoFitName(OpenXmlElement element) => element switch
    {
        A.NoAutoFit => "none",
        A.NormalAutoFit => "shrinkText",
        A.ShapeAutoFit => "resizeShape",
        _ => string.Empty,
    };

    private static string TextBodyAutoFitName(string elementLocalName) => elementLocalName switch
    {
        "noAutofit" => "none",
        "normAutofit" => "shrinkText",
        "spAutoFit" => "resizeShape",
        _ => string.Empty,
    };

    private static string TextBodyAutoFitElementName(string value) => value switch
    {
        "none" => "noAutofit",
        "shrinkText" => "normAutofit",
        "resizeShape" => "spAutoFit",
        _ => throw new CodecException("invalid_presentation_edit_target", $"Unsupported Presentation text-body AutoFit mode {value}."),
    };

    private static bool IsBareAutoFitChoice(OpenXmlElement element) => element.GetAttributes().Count == 0 && element.ChildElements.Count == 0;

    private static string ParseTextBodyAutoFitToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("none" or "shrinkText" or "resizeShape"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body AutoFit mode must be none, shrinkText, or resizeShape.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyWarpPresetToken(string value, PresentationEditOperation operation)
    {
        try
        {
            return PptxBodyPropertiesCodec.ParseTextWarpPreset(value);
        }
        catch (CodecException)
        {
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body warp preset is outside the supported DrawingML TextShapeValues vocabulary.", operation.SlidePartPath);
        }
    }

    private static string ParseTextBodyWarpAdjustmentToken(string value, PresentationEditOperation operation)
    {
        try
        {
            return PptxBodyPropertiesCodec.ParseTextWarpAdjustment(value).ToString(CultureInfo.InvariantCulture);
        }
        catch (CodecException)
        {
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body warp adjustment must use a canonical signed 32-bit integer token.", operation.SlidePartPath);
        }
    }

    private static string ParseTextBodyFlatTextZToken(string value, PresentationEditOperation operation)
    {
        try
        {
            return PptxBodyPropertiesCodec.ParseFlatTextZ(value).ToString(CultureInfo.InvariantCulture);
        }
        catch (CodecException)
        {
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body flat-text z coordinate must use a canonical signed 32-bit integer token.", operation.SlidePartPath);
        }
    }

    private static string ParseTextBodyNormalAutoFitToken(string value, string leafKind, PresentationEditOperation operation)
    {
        var maximum = leafKind == "textBodyNormalAutoFitFontScale" ? 100_000 : 13_200_000;
        var minimum = leafKind == "textBodyNormalAutoFitFontScale" ? 1_000 : 0;
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) ||
            parsed < minimum || parsed > maximum ||
            parsed.ToString(CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} {leafKind} must use a canonical thousandth-of-a-percent token.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyColumnDirectionToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("0" or "1"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body column direction must use canonical 0 or 1.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyVerticalTextToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("horizontal" or "vertical" or "vertical270"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body vertical text mode must be horizontal, vertical, or vertical270.", operation.SlidePartPath);
        return value;
    }

    private static int ParseParagraphSpacingToken(string value, string leafKind, PresentationEditOperation operation)
    {
        var points = leafKind.EndsWith("Points", StringComparison.Ordinal);
        var maximum = points ? 158_400 : 13_200_000;
        var minimum = leafKind is "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" ? 1 : 0;
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed < minimum || parsed > maximum ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} {leafKind} must use a canonical DrawingML spacing token.", operation.SlidePartPath);
        return parsed;
    }

    private static bool ValidParagraphSpacingNative(int value, bool multiplier, bool allowZero, out string token)
    {
        var maximum = multiplier ? 13_200_000 : 158_400;
        if (value < (allowZero ? 0 : 1) || value > maximum)
        {
            token = string.Empty;
            return false;
        }
        token = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    private static long ParseParagraphLayoutToken(string value, string leafKind, PresentationEditOperation operation)
    {
        const long maximum = 51_206_400;
        if (!long.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed < (leafKind == "paragraphMarginLeftEmu" ? 0 : -maximum) || parsed > maximum ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} {leafKind} must use a canonical EMU integer.", operation.SlidePartPath);
        return parsed;
    }

    private static bool TryReadBulletCharacter(A.CharacterBullet source, out string value)
    {
        value = source.Char?.Value ?? string.Empty;
        var attributes = source.GetAttributes();
        return source.ChildElements.Count == 0 && attributes.Count == 1 &&
            attributes[0].LocalName == "char" && TryBulletCharacter(value);
    }

    private static bool TryReadAutoNumber(A.AutoNumberedBullet source, out string scheme, out string? startAt)
    {
        scheme = source.Type?.InnerText ?? string.Empty;
        startAt = source.StartAt?.Value is { } start ? start.ToString(System.Globalization.CultureInfo.InvariantCulture) : null;
        var attributes = source.GetAttributes();
        var validAttributes = attributes.All(attribute => attribute.LocalName is "type" or "startAt") &&
            attributes.Count(attribute => attribute.LocalName == "type") == 1 &&
            attributes.Count(attribute => attribute.LocalName == "startAt") <= 1;
        return source.ChildElements.Count == 0 && validAttributes &&
            PptxBulletCodec.IsAutoNumberScheme(scheme) &&
            (startAt is null || ParseParagraphAutoNumberStartAtToken(startAt, null) > 0);
    }

    private static string ReadParagraphBulletStyleValue(
        A.TextParagraphPropertiesType properties,
        string kind,
        PresentationEditOperation operation)
    {
        if (kind == "paragraphBulletFontFamily")
        {
            var fonts = properties.ChildElements.OfType<A.BulletFont>().ToArray();
            if (fonts.Length != 1 || fonts[0].ChildElements.Count != 0 || fonts[0].GetAttributes().Count != 1 ||
                fonts[0].GetAttributes()[0].LocalName != "typeface" || fonts[0].Typeface?.Value is not { } family || !ValidFontFamilyToken(family))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct bullet font family.", operation.SlidePartPath);
            return family;
        }
        if (kind is "paragraphBulletColorRgb" or "paragraphBulletColorScheme")
        {
            var colors = properties.ChildElements.OfType<A.BulletColor>().ToArray();
            if (colors.Length != 1 || colors[0].GetAttributes().Count != 0 || colors[0].ChildElements.Count != 1)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct bullet color.", operation.SlidePartPath);
            var child = colors[0].FirstChild!;
            if (child.ChildElements.Count != 0 || child.GetAttributes().Count != 1 || child.GetAttributes()[0].LocalName != "val")
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has an unbounded direct bullet color.", operation.SlidePartPath);
            var value = child.GetAttributes()[0].Value ?? string.Empty;
            if (kind == "paragraphBulletColorRgb" && child is A.RgbColorModelHex)
                return PptxColor.Normalize(value);
            if (kind == "paragraphBulletColorScheme" && child is A.SchemeColor && PptxColor.TrySchemeToken(value, out var scheme))
                return scheme;
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has the wrong direct bullet color model.", operation.SlidePartPath);
        }
        var sizeKind = kind == "paragraphBulletSizePoints" ? typeof(A.BulletSizePoints) : typeof(A.BulletSizePercentage);
        var sizes = properties.ChildElements.Where(child => child.GetType() == sizeKind).ToArray();
        if (sizes.Length != 1 || sizes[0].ChildElements.Count != 0 || sizes[0].GetAttributes().Count != 1 || sizes[0].GetAttributes()[0].LocalName != "val")
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded direct bullet size.", operation.SlidePartPath);
        return ParseParagraphBulletSizeToken(sizes[0].GetAttributes()[0].Value ?? string.Empty, kind, operation)
            .ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static bool TryBulletCharacter(string? value) =>
        !string.IsNullOrEmpty(value) &&
        value.EnumerateRunes().Count() == 1 &&
        value.EnumerateRunes().All(rune => !Rune.IsControl(rune));

    private static int ParseParagraphLevelToken(string value, PresentationEditOperation operation)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 0 or > 8 || parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} paragraphLevel must use a canonical integer from 0 through 8.", operation.SlidePartPath);
        return parsed;
    }

    private static int ParseParagraphAutoNumberStartAtToken(string value, PresentationEditOperation? operation)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < 1 or > 32_767 || parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation?.OperationId ?? "(probe)"} auto-number start must be a canonical integer from 1 through 32767.", operation?.SlidePartPath);
        return parsed;
    }

    private static int ParseParagraphBulletSizeToken(string value, string kind, PresentationEditOperation operation)
    {
        var (minimum, maximum) = kind == "paragraphBulletSizePoints" ? (100, 76_800) : (25_000, 400_000);
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed < minimum || parsed > maximum || parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} {kind} must use a canonical bullet-size token.", operation.SlidePartPath);
        return parsed;
    }

    private static string EscapeXmlAttribute(string value) => value
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace("\"", "&quot;", StringComparison.Ordinal)
        .Replace("'", "&apos;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal);

    private static int ParseTextBodyColumnCountToken(string value, PresentationEditOperation operation)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed) || parsed is < 1 or > 16)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body column count must be an integer from 1 through 16.", operation.SlidePartPath);
        return parsed;
    }

    private static int ParseTextBodyColumnGapToken(string value, PresentationEditOperation operation)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed) || parsed < 0)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body column gap must be a non-negative EMU integer.", operation.SlidePartPath);
        return parsed;
    }

    private static int ParseTextBodyRotationToken(string value, PresentationEditOperation operation)
    {
        if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < -21_600_000 or > 21_600_000 ||
            parsed.ToString(System.Globalization.CultureInfo.InvariantCulture) != value)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body rotation must be a canonical integer from -21600000 through 21600000 (60000ths of a degree).", operation.SlidePartPath);
        return parsed;
    }

    private static string TextBodyVerticalOverflowName(A.TextVerticalOverflowValues? value) => value is { } current && current == A.TextVerticalOverflowValues.Overflow
        ? "overflow"
        : value is { } ellipsis && ellipsis == A.TextVerticalOverflowValues.Ellipsis
            ? "ellipsis"
            : value is { } clip && clip == A.TextVerticalOverflowValues.Clip
                ? "clip"
                : string.Empty;

    private static string ParseTextBodyVerticalOverflowToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("overflow" or "ellipsis" or "clip"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body vertical overflow must be overflow, ellipsis, or clip.", operation.SlidePartPath);
        return value;
    }

    private static string TextBodyHorizontalOverflowName(A.TextHorizontalOverflowValues? value) => value is { } current && current == A.TextHorizontalOverflowValues.Overflow
        ? "overflow"
        : value is { } clip && clip == A.TextHorizontalOverflowValues.Clip
            ? "clip"
            : string.Empty;

    private static string ParseTextBodyHorizontalOverflowToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("overflow" or "clip"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body horizontal overflow must be overflow or clip.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyUprightToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("0" or "1"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body upright must use canonical 0 or 1.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyAnchorCenterToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("0" or "1"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body anchor-center must use canonical 0 or 1.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyForceAntiAliasToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("0" or "1"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body force anti-alias must use canonical 0 or 1.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodySpaceFirstLastParagraphToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("0" or "1"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body first-last paragraph spacing must use canonical 0 or 1.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyCompatibleLineSpacingToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("0" or "1"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body compatible line spacing must use canonical 0 or 1.", operation.SlidePartPath);
        return value;
    }

    private static string ParseTextBodyFromWordArtToken(string value, PresentationEditOperation operation)
    {
        if (value is not ("0" or "1"))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body WordArt marker must use canonical 0 or 1.", operation.SlidePartPath);
        return value;
    }

    private static string ParseBoundedInsetToken(string value, PresentationEditOperation operation)
    {
        if (!long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var parsed) || parsed < 0 || parsed > int.MaxValue)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} text-body inset is outside the safe EMU range.", operation.SlidePartPath);
        return parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static PptxXmlPatch CompileTextFontSizeXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} fontSizePoints leaf is not owned by an editable text run.", operation.SlidePartPath);
        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var propertiesXml = elementXml[properties.Start..properties.End];
        var startTag = XmlTokenPattern().Matches(propertiesXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "rPr") ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} run properties tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "sz")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} run font size attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (valueGroup.Value != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw run font size does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + properties.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextGlowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} leaf is not owned by an editable text run.", operation.SlidePartPath);

        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var effectList = DirectChildRange(elementXml, properties, "rPr", "effectLst", operation);
        var glow = DirectChildRange(elementXml, effectList, "effectLst", "glow", operation);
        XmlRange target;
        string attribute;
        string expectedScalar;
        string replacement;
        if (kind == "textGlowRadiusEmu")
        {
            target = glow;
            attribute = "rad";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }
        else if (kind is "textGlowColorRgb" or "textGlowColorScheme")
        {
            var colorName = kind == "textGlowColorRgb" ? "srgbClr" : "schemeClr";
            target = DirectChildRange(elementXml, glow, "glow", colorName, operation);
            attribute = "val";
            expectedScalar = kind == "textGlowColorRgb"
                ? PptxColor.Normalize(operation.ExpectedValue)
                : PptxColor.NormalizeScheme(operation.ExpectedValue);
            replacement = kind == "textGlowColorRgb"
                ? PptxColor.Normalize(operation.Value)
                : PptxColor.NormalizeScheme(operation.Value);
        }
        else
        {
            var colors = DirectChildRanges(elementXml, glow)
                .Where(entry => entry.LocalName is "srgbClr" or "schemeClr")
                .ToArray();
            if (colors.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text glow color is missing or ambiguous.", operation.SlidePartPath);
            target = DirectChildRange(elementXml, colors[0], colors[0].LocalName, "alpha", operation);
            attribute = "val";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }

        var fragment = elementXml[target.Start..target.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == target.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text glow scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text glow scalar attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var matches = kind == "textGlowColorRgb"
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expectedScalar
            : kind == "textGlowColorScheme"
                ? PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expectedScalar
                : actual == expectedScalar;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text glow value does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + target.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextInnerShadowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} leaf is not owned by an editable text run.", operation.SlidePartPath);

        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var effectList = DirectChildRange(elementXml, properties, "rPr", "effectLst", operation);
        var innerShadow = DirectChildRange(elementXml, effectList, "effectLst", "innerShdw", operation);
        XmlRange target;
        string attribute;
        string expectedScalar;
        string replacement;
        if (kind is "textInnerShadowBlurRadiusEmu" or "textInnerShadowDistanceEmu" or "textInnerShadowDirectionDegrees")
        {
            target = innerShadow;
            attribute = kind switch
            {
                "textInnerShadowBlurRadiusEmu" => "blurRad",
                "textInnerShadowDistanceEmu" => "dist",
                _ => "dir",
            };
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }
        else if (kind is "textInnerShadowColorRgb" or "textInnerShadowColorScheme")
        {
            var colorName = kind == "textInnerShadowColorRgb" ? "srgbClr" : "schemeClr";
            target = DirectChildRange(elementXml, innerShadow, "innerShdw", colorName, operation);
            attribute = "val";
            expectedScalar = kind == "textInnerShadowColorRgb"
                ? PptxColor.Normalize(operation.ExpectedValue)
                : PptxColor.NormalizeScheme(operation.ExpectedValue);
            replacement = kind == "textInnerShadowColorRgb"
                ? PptxColor.Normalize(operation.Value)
                : PptxColor.NormalizeScheme(operation.Value);
        }
        else
        {
            var colors = DirectChildRanges(elementXml, innerShadow)
                .Where(entry => entry.LocalName is "srgbClr" or "schemeClr")
                .ToArray();
            if (colors.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text inner-shadow color is missing or ambiguous.", operation.SlidePartPath);
            target = DirectChildRange(elementXml, colors[0], colors[0].LocalName, "alpha", operation);
            attribute = "val";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }

        var fragment = elementXml[target.Start..target.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == target.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text inner-shadow scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} text inner-shadow scalar attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var matches = kind == "textInnerShadowColorRgb"
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expectedScalar
            : kind == "textInnerShadowColorScheme"
                ? PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expectedScalar
                : actual == expectedScalar;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text inner-shadow value does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + target.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextReflectionXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} leaf is not owned by an editable text run.", operation.SlidePartPath);

        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var effectList = DirectChildRange(elementXml, properties, "rPr", "effectLst", operation);
        var reflection = DirectChildRange(elementXml, effectList, "effectLst", "reflection", operation);
        var attribute = kind switch
        {
            "textReflectionBlurRadiusEmu" => "blurRad",
            "textReflectionStartOpacityThousandthPercent" => "stA",
            "textReflectionEndOpacityThousandthPercent" => "endA",
            "textReflectionDistanceEmu" => "dist",
            "textReflectionDirectionDegrees" => "dir",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported text reflection leaf.", operation.SlidePartPath),
        };
        var fragment = elementXml[reflection.Start..reflection.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "reflection")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} reflection scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} reflection scalar attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        if (actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text reflection value does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + reflection.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextSoftEdgeXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textSoftEdgeRadiusEmu leaf is not owned by an editable text run.", operation.SlidePartPath);

        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var effectList = DirectChildRange(elementXml, properties, "rPr", "effectLst", operation);
        var softEdge = DirectChildRange(elementXml, effectList, "effectLst", "softEdge", operation);
        var fragment = elementXml[softEdge.Start..softEdge.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "softEdge")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} soft-edge scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "rad")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} soft-edge radius attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (System.Net.WebUtility.HtmlDecode(valueGroup.Value) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw text soft-edge radius does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + softEdge.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultSoftEdgeXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var softEdge = DirectChildRange(elementXml, effectList, "effectLst", "softEdge", operation);
        var fragment = elementXml[softEdge.Start..softEdge.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "softEdge")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text soft-edge scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "rad")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text soft-edge radius attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (System.Net.WebUtility.HtmlDecode(valueGroup.Value) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text soft-edge radius does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + softEdge.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultGlowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var glow = DirectChildRange(elementXml, effectList, "effectLst", "glow", operation);
        XmlRange target;
        string attribute;
        string expectedScalar;
        string replacement;
        if (kind == "textDefaultGlowRadiusEmu")
        {
            target = glow;
            attribute = "rad";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }
        else if (kind is "textDefaultGlowColorRgb" or "textDefaultGlowColorScheme")
        {
            var colorName = kind == "textDefaultGlowColorRgb" ? "srgbClr" : "schemeClr";
            target = DirectChildRange(elementXml, glow, "glow", colorName, operation);
            attribute = "val";
            expectedScalar = kind == "textDefaultGlowColorRgb"
                ? PptxColor.Normalize(operation.ExpectedValue)
                : PptxColor.NormalizeScheme(operation.ExpectedValue);
            replacement = kind == "textDefaultGlowColorRgb"
                ? PptxColor.Normalize(operation.Value)
                : PptxColor.NormalizeScheme(operation.Value);
        }
        else if (kind == "textDefaultGlowOpacityThousandthPercent")
        {
            var colors = DirectChildRanges(elementXml, glow)
                .Where(range => range.LocalName is "srgbClr" or "schemeClr")
                .ToArray();
            if (colors.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text glow color is missing or ambiguous.", operation.SlidePartPath);
            var color = colors[0];
            target = DirectChildRange(elementXml, color, color.LocalName, "alpha", operation);
            attribute = "val";
            expectedScalar = operation.ExpectedValue;
            replacement = operation.Value;
        }
        else
        {
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph default-text glow leaf.", operation.SlidePartPath);
        }
        var fragment = elementXml[target.Start..target.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == target.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text glow scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text glow {attribute} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var matches = kind == "textDefaultGlowColorRgb"
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expectedScalar
            : kind == "textDefaultGlowColorScheme"
                ? PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expectedScalar
                : actual == expectedScalar;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text glow value does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + target.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultShadowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var outerShadow = DirectChildRange(elementXml, effectList, "effectLst", "outerShdw", operation);
        var fragment = elementXml[outerShadow.Start..outerShadow.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "outerShdw")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow scalar tag was not found.", operation.SlidePartPath);
        var attribute = LeafKind(operation) switch
        {
            "textDefaultShadowBlurRadiusEmu" => "blurRad",
            "textDefaultShadowDistanceEmu" => "dist",
            "textDefaultShadowDirectionDegrees" => "dir",
            "textDefaultShadowAlignment" => "algn",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph default-text outer-shadow leaf.", operation.SlidePartPath),
        };
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow {attribute} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (System.Net.WebUtility.HtmlDecode(valueGroup.Value) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text outer-shadow {attribute} does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + outerShadow.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultShadowColorXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var outerShadow = DirectChildRange(elementXml, effectList, "effectLst", "outerShdw", operation);
        var kind = LeafKind(operation);
        var colorName = kind == "textDefaultShadowColorRgb" ? "srgbClr" : "schemeClr";
        var color = DirectChildRange(elementXml, outerShadow, "outerShdw", colorName, operation);
        var fragment = elementXml[color.Start..color.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == colorName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow color tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow color value attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var isRgb = kind == "textDefaultShadowColorRgb";
        var expected = isRgb ? PptxColor.Normalize(operation.ExpectedValue) : PptxColor.NormalizeScheme(operation.ExpectedValue);
        var replacement = isRgb ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
        var matches = isRgb
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expected
            : PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expected;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text outer-shadow color does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + color.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultShadowAlphaXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var outerShadow = DirectChildRange(elementXml, effectList, "effectLst", "outerShdw", operation);
        var colors = DirectChildRanges(elementXml, outerShadow)
            .Where(range => range.LocalName is "srgbClr" or "schemeClr")
            .ToArray();
        if (colors.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow color is missing or ambiguous.", operation.SlidePartPath);
        var color = colors[0];
        var alpha = DirectChildRange(elementXml, color, color.LocalName, "alpha", operation);
        var fragment = elementXml[alpha.Start..alpha.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "alpha")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow alpha tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow alpha value attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        if (!ValidOpacityToken(actual) || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text outer-shadow alpha does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + alpha.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultShadowRotateWithShapeXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var outerShadow = DirectChildRange(elementXml, effectList, "effectLst", "outerShdw", operation);
        var fragment = elementXml[outerShadow.Start..outerShadow.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "outerShdw")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "rotWithShape")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text outer-shadow rotWithShape attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        if (!ValidBooleanToken(actual) || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text outer-shadow rotWithShape does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + outerShadow.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultInnerShadowXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var innerShadow = DirectChildRange(elementXml, effectList, "effectLst", "innerShdw", operation);
        var attribute = LeafKind(operation) switch
        {
            "textDefaultInnerShadowBlurRadiusEmu" => "blurRad",
            "textDefaultInnerShadowDistanceEmu" => "dist",
            "textDefaultInnerShadowDirectionDegrees" => "dir",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph default-text inner-shadow leaf.", operation.SlidePartPath),
        };
        var fragment = elementXml[innerShadow.Start..innerShadow.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "innerShdw")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text inner-shadow scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text inner-shadow {attribute} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (System.Net.WebUtility.HtmlDecode(valueGroup.Value) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text inner-shadow {attribute} does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + innerShadow.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultInnerShadowColorXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var innerShadow = DirectChildRange(elementXml, effectList, "effectLst", "innerShdw", operation);
        var kind = LeafKind(operation);
        var colorName = kind == "textDefaultInnerShadowColorRgb" ? "srgbClr" : "schemeClr";
        var color = DirectChildRange(elementXml, innerShadow, "innerShdw", colorName, operation);
        var fragment = elementXml[color.Start..color.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == colorName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text inner-shadow color tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text inner-shadow color value attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var expected = kind == "textDefaultInnerShadowColorRgb"
            ? PptxColor.Normalize(operation.ExpectedValue)
            : PptxColor.NormalizeScheme(operation.ExpectedValue);
        var replacement = kind == "textDefaultInnerShadowColorRgb"
            ? PptxColor.Normalize(operation.Value)
            : PptxColor.NormalizeScheme(operation.Value);
        var matches = kind == "textDefaultInnerShadowColorRgb"
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expected
            : PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expected;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text inner-shadow color does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + color.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultInnerShadowAlphaXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var innerShadow = DirectChildRange(elementXml, effectList, "effectLst", "innerShdw", operation);
        var colors = DirectChildRanges(elementXml, innerShadow)
            .Where(range => range.LocalName is "srgbClr" or "schemeClr")
            .ToArray();
        if (colors.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text inner-shadow color is missing or ambiguous.", operation.SlidePartPath);
        var color = colors[0];
        var alpha = DirectChildRange(elementXml, color, color.LocalName, "alpha", operation);
        var fragment = elementXml[alpha.Start..alpha.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "alpha")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text inner-shadow alpha tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text inner-shadow alpha value attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        if (!ValidOpacityToken(actual) || actual != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text inner-shadow alpha does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + alpha.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextDefaultReflectionXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var shapeRange = new XmlRange(0, elementXml.Length, "sp");
        var txBody = DirectChildRange(elementXml, shapeRange, "sp", "txBody", operation);
        var paragraphs = DirectChildRanges(elementXml, txBody)
            .Where(range => range.LocalName == "p")
            .ToArray();
        if (operation.NativeLeafIndex >= (uint)paragraphs.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text index is out of range.", operation.SlidePartPath);
        var paragraph = paragraphs[operation.NativeLeafIndex];
        var pPr = DirectChildRange(elementXml, paragraph, "p", "pPr", operation);
        var defaultRunProperties = DirectChildRange(elementXml, pPr, "pPr", "defRPr", operation);
        var effectList = DirectChildRange(elementXml, defaultRunProperties, "defRPr", "effectLst", operation);
        var reflection = DirectChildRange(elementXml, effectList, "effectLst", "reflection", operation);
        var attribute = LeafKind(operation) switch
        {
            "textDefaultReflectionBlurRadiusEmu" => "blurRad",
            "textDefaultReflectionDistanceEmu" => "dist",
            "textDefaultReflectionStartOpacityThousandthPercent" => "stA",
            "textDefaultReflectionEndOpacityThousandthPercent" => "endA",
            "textDefaultReflectionDirectionDegrees" => "dir",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported paragraph default-text reflection leaf.", operation.SlidePartPath),
        };
        var fragment = elementXml[reflection.Start..reflection.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "reflection")
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} paragraph default-text reflection scalar tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} paragraph default-text reflection {attribute} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (System.Net.WebUtility.HtmlDecode(valueGroup.Value) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw paragraph default-text reflection {attribute} does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + reflection.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextFontFamilyXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {LeafKind(operation)} leaf is not owned by an editable text run.", operation.SlidePartPath);
        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var fontName = LeafKind(operation) switch
        {
            "fontFamily" => "latin",
            "fontFamilyEastAsia" => "ea",
            "fontFamilyComplexScript" => "cs",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported font-family leaf.", operation.SlidePartPath),
        };
        var fonts = DirectChildRanges(elementXml, properties).Where(entry => entry.LocalName == fontName).ToArray();
        if (fonts.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} run {fontName} font is missing or ambiguous.", operation.SlidePartPath);
        var font = fonts[0];
        var fragment = elementXml[font.Start..font.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == fontName) ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} run {fontName} font tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "typeface")
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} run {fontName} typeface attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        if (System.Net.WebUtility.HtmlDecode(valueGroup.Value) != operation.ExpectedValue)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw run font family does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + font.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, EscapeAttribute(operation.Value), proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextFontBooleanXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {LeafKind(operation)} leaf is not owned by an editable text run.", operation.SlidePartPath);
        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var propertiesXml = elementXml[properties.Start..properties.End];
        var startTag = XmlTokenPattern().Matches(propertiesXml).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "rPr") ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} run properties tag was not found.", operation.SlidePartPath);
        var leafKind = LeafKind(operation);
        var attribute = leafKind switch
        {
            "fontBold" => "b",
            "fontItalic" => "i",
            "fontUnderline" => "u",
            "fontStrike" => "strike",
            "fontKerningPoints" => "kern",
            "fontBaselinePercent" => "baseline",
            "fontSpacingPoints" => "spc",
            "fontCaps" => "cap",
            "fontLanguage" => "lang",
            _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an unsupported text decoration leaf.", operation.SlidePartPath),
        };
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} run {attribute} attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actualValue = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var matches = leafKind is "fontBold" or "fontItalic"
            ? TryCanonicalBoolean(actualValue, out var actualBoolean) && actualBoolean == operation.ExpectedValue
            : leafKind == "fontUnderline"
                ? PptxTextDecoration.IsUnderlineToken(actualValue) && actualValue == operation.ExpectedValue
                : leafKind == "fontStrike"
                    ? PptxTextDecoration.IsStrikeToken(actualValue) && actualValue == operation.ExpectedValue
                    : leafKind == "fontKerningPoints"
                        ? PptxTextDecoration.IsKerningToken(actualValue) && actualValue == operation.ExpectedValue
                        : leafKind == "fontBaselinePercent"
                            ? PptxTextDecoration.IsBaselineToken(actualValue) && actualValue == operation.ExpectedValue
                            : leafKind == "fontSpacingPoints"
                                ? PptxTextDecoration.IsSpacingToken(actualValue) && actualValue == operation.ExpectedValue
                                : leafKind == "fontCaps"
                                    ? PptxTextDecoration.IsCapsToken(actualValue) && actualValue == operation.ExpectedValue
                                    : PptxLanguageTag.IsValid(actualValue) && actualValue == operation.ExpectedValue;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw run font style does not match the expected value.", operation.SlidePartPath);
        var start = elementRange.Start + properties.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileTextFontColorXmlPatch(
        string xml,
        XmlRange elementRange,
        PptxEditPlanProof proof,
        IReadOnlySet<string> drawingPrefixes)
    {
        var operation = proof.Operation;
        var elementXml = xml[elementRange.Start..elementRange.End];
        var leaves = TextLeafPattern().Matches(elementXml)
            .Where(IsNonSelfClosingTextLeaf)
            .Where(match => drawingPrefixes.Contains(match.Groups["prefix"].Value))
            .ToArray();
        if (operation.TextLeafIndex >= (uint)leaves.Length)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw text-leaf index is out of range.", operation.SlidePartPath);
        var leaf = leaves[operation.TextLeafIndex];
        var run = TextRunPattern().Matches(elementXml).Cast<Match>()
            .SingleOrDefault(match => match.Index <= leaf.Index && leaf.Index < match.Index + match.Length);
        if (run is null)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {LeafKind(operation)} leaf is not owned by an editable text run.", operation.SlidePartPath);
        var runRange = new XmlRange(run.Index, run.Index + run.Length, "r");
        var properties = DirectChildRange(elementXml, runRange, "r", "rPr", operation);
        var leafKind = LeafKind(operation);
        var isHighlight = leafKind is "fontHighlightRgb" or "fontHighlightScheme";
        var fill = DirectChildRange(elementXml, properties, "rPr", isHighlight ? "highlight" : "solidFill", operation);
        var colorName = leafKind is "fontColorRgb" or "fontHighlightRgb" ? "srgbClr" : "schemeClr";
        var fillChildren = DirectChildRanges(elementXml, fill);
        if (fillChildren.Count != 1 || fillChildren[0].LocalName != colorName)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} run {(isHighlight ? "font highlight" : "font color")} requires one direct {colorName} child.", operation.SlidePartPath);
        var color = fillChildren[0];
        if (DirectChildRanges(elementXml, color).Count != 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} run {(isHighlight ? "font highlight" : "font color")} has unsupported color effects.", operation.SlidePartPath);
        var fragment = elementXml[color.Start..color.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == colorName) ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} run {(isHighlight ? "font highlight" : "font color")} tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        if (attributes.Length != 1 || XmlAttributePattern().Matches(startTag.Value).Cast<Match>().Any(match => LocalAttributeName(match.Groups["name"].Value) != "val"))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} run {(isHighlight ? "font highlight" : "font color")} value is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var matches = leafKind is "fontColorRgb" or "fontHighlightRgb"
            ? actual.Equals(operation.ExpectedValue, StringComparison.OrdinalIgnoreCase) && IsRgbToken(actual)
            : PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == PptxColor.NormalizeScheme(operation.ExpectedValue);
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw run {(isHighlight ? "font highlight" : "font color")} does not match the expected value.", operation.SlidePartPath);
        var replacement = leafKind is "fontColorRgb" or "fontHighlightRgb" ? PptxColor.Normalize(operation.Value) : PptxColor.NormalizeScheme(operation.Value);
        var start = elementRange.Start + color.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileNativeJoinXmlPatch(string xml, XmlRange outlineOrJoin, PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var join = outlineOrJoin.LocalName == "ln"
            ? DirectChildRanges(xml, outlineOrJoin)
                .Where(entry => entry.LocalName is "round" or "bevel" or "miter")
                .ToArray() switch
            {
                { Length: 1 } matches => matches[0],
                _ => throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one bare line-join element.", operation.SlidePartPath),
            }
            : outlineOrJoin;
        var fragment = xml[join.Start..join.End];
        var opening = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == join.LocalName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} line-join tag was not found.", operation.SlidePartPath);
        if (!opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) ||
            XmlAttributePattern().Matches(opening.Value).Count != 0 ||
            DirectChildRanges(xml, join).Count != 0 ||
            !PptxLineStyleCodec.TryJoinToken(operation.ExpectedValue, out var expectedToken) ||
            !PptxLineStyleCodec.TryJoinToken(operation.Value, out var replacementToken) ||
            !string.Equals(join.LocalName, expectedToken, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw line join does not match the expected value.", operation.SlidePartPath);
        var name = XmlLocalNamePattern().Match(opening.Value).Groups["name"];
        if (!name.Success) throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} line-join tag name is invalid.", operation.SlidePartPath);
        var start = join.Start + opening.Index + name.Index;
        return new PptxXmlPatch(operation, start, start + name.Length, replacementToken, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileNativeArrowXmlPatch(string xml, XmlRange outlineOrArrow, PptxEditPlanProof proof, string endpointName)
    {
        var operation = proof.Operation;
        var arrow = outlineOrArrow.LocalName == "ln"
            ? DirectChildRanges(xml, outlineOrArrow).Where(entry => entry.LocalName == endpointName).ToArray() switch
            {
                { Length: 1 } matches => matches[0],
                _ => throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one explicit {endpointName} element.", operation.SlidePartPath),
            }
            : outlineOrArrow;
        if (!string.Equals(arrow.LocalName, endpointName, StringComparison.Ordinal))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} arrow endpoint does not match {endpointName}.", operation.SlidePartPath);

        var fragment = xml[arrow.Start..arrow.End];
        var opening = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == endpointName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} arrow endpoint tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(opening.Value).Cast<Match>().ToArray();
        var typeAttributes = attributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "type").ToArray();
        var invalidAttribute = attributes.Any(attribute => LocalAttributeName(attribute.Groups["name"].Value) is not ("type" or "w" or "len"));
        var invalidSize = attributes
            .Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) is "w" or "len")
            .Any(attribute => !PptxLineStyleCodec.TryArrowSizeToken(attribute.Groups["value"].Value));
        if (!opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) ||
            DirectChildRanges(xml, arrow).Count != 0 ||
            typeAttributes.Length != 1 ||
            attributes.Count(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "w") > 1 ||
            attributes.Count(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "len") > 1 ||
            invalidAttribute || invalidSize ||
            !PptxLineStyleCodec.TryArrowTypeToken(operation.ExpectedValue, out var expectedToken) ||
            !PptxLineStyleCodec.TryArrowTypeToken(operation.Value, out var replacementToken) ||
            !string.Equals(typeAttributes[0].Groups["value"].Value, expectedToken, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw arrow endpoint does not match the expected value.", operation.SlidePartPath);

        var valueGroup = typeAttributes[0].Groups["value"];
        var start = arrow.Start + opening.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacementToken, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileNativeArrowSizeXmlPatch(
        string xml,
        XmlRange outlineOrArrow,
        PptxEditPlanProof proof,
        string endpointName,
        string attributeName)
    {
        var operation = proof.Operation;
        var arrow = outlineOrArrow.LocalName == "ln"
            ? DirectChildRanges(xml, outlineOrArrow).Where(entry => entry.LocalName == endpointName).ToArray() switch
            {
                { Length: 1 } matches => matches[0],
                _ => throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one explicit {endpointName} element.", operation.SlidePartPath),
            }
            : outlineOrArrow;
        if (!string.Equals(arrow.LocalName, endpointName, StringComparison.Ordinal))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} arrow endpoint does not match {endpointName}.", operation.SlidePartPath);

        var fragment = xml[arrow.Start..arrow.End];
        var opening = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == endpointName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} arrow endpoint tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(opening.Value).Cast<Match>().ToArray();
        var typeAttributes = attributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "type").ToArray();
        var sizeAttributes = attributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == attributeName).ToArray();
        var invalidAttribute = attributes.Any(attribute => LocalAttributeName(attribute.Groups["name"].Value) is not ("type" or "w" or "len"));
        var invalidSize = attributes
            .Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) is "w" or "len")
            .Any(attribute => !PptxLineStyleCodec.TryArrowSizeToken(attribute.Groups["value"].Value));
        if (!opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) ||
            DirectChildRanges(xml, arrow).Count != 0 ||
            typeAttributes.Length != 1 || sizeAttributes.Length != 1 ||
            attributes.Count(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "w") > 1 ||
            attributes.Count(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "len") > 1 ||
            invalidAttribute || invalidSize ||
            !PptxLineStyleCodec.TryArrowTypeToken(typeAttributes[0].Groups["value"].Value, out var arrowType) ||
            arrowType == "none" ||
            !PptxLineStyleCodec.TryArrowSizeToken(operation.ExpectedValue) ||
            !PptxLineStyleCodec.TryArrowSizeToken(operation.Value) ||
            !string.Equals(sizeAttributes[0].Groups["value"].Value, operation.ExpectedValue, StringComparison.Ordinal))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw arrow size does not match the expected value.", operation.SlidePartPath);

        var valueGroup = sizeAttributes[0].Groups["value"];
        var start = arrow.Start + opening.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, operation.Value, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileNativeConnectorStyleColorXmlPatch(
        string xml,
        XmlRange connectorRange,
        PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var kind = LeafKind(operation);
        var style = DirectChildRange(xml, connectorRange, "cxnSp", "style", operation);
        var lineReference = DirectChildRange(xml, style, "style", "lnRef", operation);
        var colors = DirectChildRanges(xml, lineReference)
            .Where(entry => entry.LocalName is "schemeClr" or "srgbClr")
            .ToArray();
        var expectedName = kind == "lineRgb" ? "srgbClr" : "schemeClr";
        if (kind is not ("lineRgb" or "lineScheme") || colors.Length != 1 || colors[0].LocalName != expectedName ||
            DirectChildRanges(xml, colors[0]).Count != 0)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} connector style line color is missing or ambiguous.", operation.SlidePartPath);

        var color = colors[0];
        var fragment = xml[color.Start..color.End];
        var opening = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == expectedName)
            ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} connector style color tag was not found.", operation.SlidePartPath);
        var attributes = XmlAttributePattern().Matches(opening.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        if (attributes.Length != 1 || XmlAttributePattern().Matches(opening.Value).Cast<Match>()
                .Any(match => LocalAttributeName(match.Groups["name"].Value) != "val"))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} connector style color value is missing or ambiguous.", operation.SlidePartPath);

        var valueGroup = attributes[0].Groups["value"];
        var actual = System.Net.WebUtility.HtmlDecode(valueGroup.Value);
        var expected = kind == "lineRgb"
            ? PptxColor.Normalize(operation.ExpectedValue)
            : PptxColor.NormalizeScheme(operation.ExpectedValue);
        var replacement = kind == "lineRgb"
            ? PptxColor.Normalize(operation.Value)
            : PptxColor.NormalizeScheme(operation.Value);
        var matches = kind == "lineRgb"
            ? IsRgbToken(actual) && PptxColor.Normalize(actual) == expected
            : PptxColor.TrySchemeToken(actual, out var actualScheme) && actualScheme == expected;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} connector style color does not match the expected value.", operation.SlidePartPath);
        var start = color.Start + opening.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static PptxXmlPatch CompileNativeStyleXmlPatch(string xml, XmlRange groupRange, PptxEditPlanProof proof)
    {
        var operation = proof.Operation;
        var styles = NativeStyleXmlLeaves(xml, groupRange);
        if (operation.NativeLeafIndex >= (uint)styles.Count)
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} native style-leaf index is out of range.", operation.SlidePartPath);
        var style = styles[(int)operation.NativeLeafIndex];
        if (style.Kind != LeafKind(operation))
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} native style leaf kind changed after planning.", operation.SlidePartPath);
        if (style.Kind == "lineJoin") return CompileNativeJoinXmlPatch(xml, style.Range, proof);
        if (style.Kind == "lineStartArrow") return CompileNativeArrowXmlPatch(xml, style.Range, proof, "headEnd");
        if (style.Kind == "lineEndArrow") return CompileNativeArrowXmlPatch(xml, style.Range, proof, "tailEnd");
        var fragment = xml[style.Range.Start..style.Range.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == style.Range.LocalName) ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} native style tag was not found.", operation.SlidePartPath);
        var attributeName = style.Kind switch
        {
            "lineWidthEmu" => "w",
            "lineCap" => "cap",
            _ => "val",
        };
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attributeName)
            .ToArray();
        if (attributes.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} native style attribute is missing or ambiguous.", operation.SlidePartPath);
        var valueGroup = attributes[0].Groups["value"];
        var expectedRaw = style.Kind switch
        {
            "lineStyle" => PptxLineStyleCodec.TryPresetDashToken(operation.ExpectedValue, out var expectedStyleToken) ? expectedStyleToken : string.Empty,
            "lineCap" => PptxLineStyleCodec.TryCapToken(operation.ExpectedValue, out var expectedCapToken) ? expectedCapToken : string.Empty,
            _ => operation.ExpectedValue,
        };
        var replacement = style.Kind switch
        {
            "lineStyle" => PptxLineStyleCodec.TryPresetDashToken(operation.Value, out var requestedStyleToken) ? requestedStyleToken : string.Empty,
            "lineCap" => PptxLineStyleCodec.TryCapToken(operation.Value, out var requestedCapToken) ? requestedCapToken : string.Empty,
            _ => operation.Value,
        };
        if (style.Kind is "lineStyle" or "lineCap"
            ? !string.Equals(valueGroup.Value, expectedRaw, StringComparison.Ordinal)
            : !LeafValuesEqual(valueGroup.Value, expectedRaw, style.Kind))
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw native style value does not match the expected value.", operation.SlidePartPath);
        var start = style.Range.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
    }

    private static XmlRange DirectChildRange(
        string xml,
        XmlRange parent,
        string parentLocalName,
        string childLocalName,
        PresentationEditOperation operation)
    {
        var fragment = xml[parent.Start..parent.End];
        var children = ShapeElementRanges(
                fragment,
                parentLocalName,
                includeGroupProperties: parentLocalName == "grpSp" && childLocalName == "grpSpPr")
            .Where(child => child.LocalName == childLocalName)
            .ToArray();
        if (children.Length != 1)
            throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one direct {childLocalName} child under {parentLocalName}.", operation.SlidePartPath);
        var child = children[0];
        return new XmlRange(parent.Start + child.Start, parent.Start + child.End, child.LocalName);
    }

    private static IReadOnlyList<XmlRange> DirectChildRanges(string xml, XmlRange parent)
    {
        var fragment = xml[parent.Start..parent.End];
        var opening = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal));
        if (opening is null || opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal)) return Array.Empty<XmlRange>();
        return ShapeElementRanges(fragment, parent.LocalName)
            .Select(child => new XmlRange(parent.Start + child.Start, parent.Start + child.End, child.LocalName))
            .ToArray();
    }

    private sealed record NativeStyleXmlRange(string Kind, XmlRange Range);

    private static IReadOnlyList<NativeStyleXmlRange> NativeStyleXmlLeaves(string xml, XmlRange groupRange)
    {
        var fillLeaves = new List<NativeStyleXmlRange>();
        var lineLeaves = new List<NativeStyleXmlRange>();
        var lineWidthLeaves = new List<NativeStyleXmlRange>();
        var lineStyleLeaves = new List<NativeStyleXmlRange>();
        var lineCapLeaves = new List<NativeStyleXmlRange>();
        var lineJoinLeaves = new List<NativeStyleXmlRange>();
        var lineStartArrowLeaves = new List<NativeStyleXmlRange>();
        var lineEndArrowLeaves = new List<NativeStyleXmlRange>();
        var fillNames = new HashSet<string>(StringComparer.Ordinal) { "noFill", "solidFill", "gradFill", "blipFill", "pattFill" };

        void VisitGroup(XmlRange current)
        {
            foreach (var child in DirectChildRanges(xml, current))
            {
                if (child.LocalName == "grpSp")
                {
                    VisitGroup(child);
                    continue;
                }
                if (child.LocalName != "sp") continue;
                var properties = DirectChildRanges(xml, child).Where(entry => entry.LocalName == "spPr").ToArray();
                if (properties.Length != 1) continue;
                var fills = DirectChildRanges(xml, properties[0]).Where(entry => fillNames.Contains(entry.LocalName)).ToArray();
                if (fills.Length == 1 && fills[0].LocalName == "solidFill" && TryNativeStyleXmlColor(xml, fills[0], "fill", out var fill))
                    fillLeaves.Add(fill);
                var outlines = DirectChildRanges(xml, properties[0]).Where(entry => entry.LocalName == "ln").ToArray();
                if (outlines.Length != 1) continue;
                if (TryNativeStyleXmlWidth(xml, outlines[0], out var width))
                    lineWidthLeaves.Add(width);
                if (TryNativeStyleXmlDash(xml, outlines[0], out var style))
                    lineStyleLeaves.Add(style);
                if (TryNativeStyleXmlCap(xml, outlines[0], out var cap))
                    lineCapLeaves.Add(cap);
                if (TryNativeStyleXmlJoin(xml, outlines[0], out var join))
                    lineJoinLeaves.Add(join);
                var lineFills = DirectChildRanges(xml, outlines[0]).Where(entry => fillNames.Contains(entry.LocalName)).ToArray();
                if (lineFills.Length == 1 && lineFills[0].LocalName == "solidFill" && TryNativeStyleXmlColor(xml, lineFills[0], "line", out var line))
                    lineLeaves.Add(line);
                if (TryNativeStyleXmlArrow(xml, outlines[0], "headEnd", "lineStartArrow", out var startArrow))
                    lineStartArrowLeaves.Add(startArrow);
                if (TryNativeStyleXmlArrow(xml, outlines[0], "tailEnd", "lineEndArrow", out var endArrow))
                    lineEndArrowLeaves.Add(endArrow);
                if (fillLeaves.Count + lineLeaves.Count + lineWidthLeaves.Count + lineStyleLeaves.Count + lineCapLeaves.Count + lineJoinLeaves.Count + lineStartArrowLeaves.Count + lineEndArrowLeaves.Count > 4_096)
                    throw new CodecException("presentation_item_budget_exceeded", "PPTX native opaque-group style leaves exceed the bounded style profile.");
            }
        }

        VisitGroup(groupRange);
        return fillLeaves.Concat(lineLeaves).Concat(lineWidthLeaves).Concat(lineStyleLeaves).Concat(lineCapLeaves).Concat(lineJoinLeaves).Concat(lineStartArrowLeaves).Concat(lineEndArrowLeaves).ToArray();
    }

    private static bool TryNativeStyleXmlColor(string xml, XmlRange solidFill, string prefix, out NativeStyleXmlRange color)
    {
        color = null!;
        var colors = DirectChildRanges(xml, solidFill).Where(entry => entry.LocalName is "schemeClr" or "srgbClr").ToArray();
        if (colors.Length != 1 || DirectChildRanges(xml, colors[0]).Count != 0) return false;
        var value = NativeStyleXmlAttribute(xml, colors[0]);
        if (value is null) return false;
        if (colors[0].LocalName == "schemeClr" && PptxColor.TrySchemeToken(value, out _))
        {
            color = new NativeStyleXmlRange($"{prefix}Scheme", colors[0]);
            return true;
        }
        if (colors[0].LocalName == "srgbClr" && value.Length == 6 && value.All(Uri.IsHexDigit))
        {
            color = new NativeStyleXmlRange($"{prefix}Rgb", colors[0]);
            return true;
        }
        return false;
    }

    private static bool TryNativeStyleXmlWidth(string xml, XmlRange outline, out NativeStyleXmlRange width)
    {
        width = null!;
        var fragment = xml[outline.Start..outline.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "ln");
        if (startTag is null) return false;
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "w")
            .ToArray();
        if (attributes.Length != 1 || !ulong.TryParse(
                attributes[0].Groups["value"].Value,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out var parsed) || parsed > 20_116_800)
            return false;
        var canonical = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (!string.Equals(attributes[0].Groups["value"].Value, canonical, StringComparison.Ordinal)) return false;
        width = new NativeStyleXmlRange("lineWidthEmu", new XmlRange(
            outline.Start + startTag.Index,
            outline.Start + startTag.Index + startTag.Length,
            "ln"));
        return true;
    }

    private static bool TryNativeStyleXmlDash(string xml, XmlRange outline, out NativeStyleXmlRange style)
    {
        style = null!;
        var dashes = DirectChildRanges(xml, outline).Where(entry => entry.LocalName == "prstDash").ToArray();
        if (dashes.Length != 1 || !PptxLineStyleCodec.TryReadPresetDashValue(NativeStyleXmlAttribute(xml, dashes[0]), out _)) return false;
        var fills = DirectChildRanges(xml, outline).Where(entry => entry.LocalName is "noFill" or "solidFill").ToArray();
        if (fills.Length != 1 || fills[0].LocalName != "solidFill" || !TryNativeStyleXmlColor(xml, fills[0], "line", out _)) return false;
        style = new NativeStyleXmlRange("lineStyle", dashes[0]);
        return true;
    }

    private static bool TryNativeStyleXmlCap(string xml, XmlRange outline, out NativeStyleXmlRange cap)
    {
        cap = null!;
        var fragment = xml[outline.Start..outline.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == "ln");
        if (startTag is null) return false;
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "cap")
            .ToArray();
        var outlineAttributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>().ToArray();
        if (attributes.Length != 1 ||
            outlineAttributes.Any(attribute => LocalAttributeName(attribute.Groups["name"].Value) is not ("w" or "cap" or "cmpd" or "algn")) ||
            !PptxLineStyleCodec.TryReadCapValue(attributes[0].Groups["value"].Value, out _)) return false;
        var fills = DirectChildRanges(xml, outline).Where(entry => entry.LocalName is "noFill" or "solidFill").ToArray();
        if (fills.Length != 1 || fills[0].LocalName != "solidFill" || !TryNativeStyleXmlColor(xml, fills[0], "line", out _)) return false;
        cap = new NativeStyleXmlRange("lineCap", new XmlRange(
            outline.Start + startTag.Index,
            outline.Start + startTag.Index + startTag.Length,
            "ln"));
        return true;
    }

    private static bool TryNativeStyleXmlJoin(string xml, XmlRange outline, out NativeStyleXmlRange join)
    {
        join = null!;
        var joins = DirectChildRanges(xml, outline)
            .Where(entry => entry.LocalName is "round" or "bevel" or "miter")
            .ToArray();
        if (joins.Length != 1) return false;
        var candidate = joins[0];
        var fragment = xml[candidate.Start..candidate.End];
        var opening = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == candidate.LocalName);
        if (opening is null || !opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) ||
            XmlAttributePattern().Matches(opening.Value).Count != 0 ||
            DirectChildRanges(xml, candidate).Count != 0 ||
            !TryNativeStyleXmlLinePaint(xml, outline)) return false;
        join = new NativeStyleXmlRange("lineJoin", candidate);
        return true;
    }

    private static bool TryNativeStyleXmlArrow(string xml, XmlRange outline, string endpointName, string kind, out NativeStyleXmlRange arrow)
    {
        arrow = null!;
        var endpoints = DirectChildRanges(xml, outline).Where(entry => entry.LocalName == endpointName).ToArray();
        if (endpoints.Length != 1 || !TryNativeStyleXmlLinePaint(xml, outline)) return false;
        var candidate = endpoints[0];
        var fragment = xml[candidate.Start..candidate.End];
        var opening = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == endpointName);
        if (opening is null || !opening.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) || DirectChildRanges(xml, candidate).Count != 0) return false;
        var attributes = XmlAttributePattern().Matches(opening.Value).Cast<Match>().ToArray();
        var typeAttributes = attributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "type").ToArray();
        if (typeAttributes.Length != 1 ||
            attributes.Any(attribute => LocalAttributeName(attribute.Groups["name"].Value) is not ("type" or "w" or "len")) ||
            attributes.Count(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "w") > 1 ||
            attributes.Count(attribute => LocalAttributeName(attribute.Groups["name"].Value) == "len") > 1 ||
            attributes.Where(attribute => LocalAttributeName(attribute.Groups["name"].Value) is "w" or "len")
                .Any(attribute => !PptxLineStyleCodec.TryArrowSizeToken(attribute.Groups["value"].Value)) ||
            !PptxLineStyleCodec.TryArrowTypeToken(typeAttributes[0].Groups["value"].Value, out _)) return false;
        arrow = new NativeStyleXmlRange(kind, candidate);
        return true;
    }

    private static bool TryNativeStyleXmlLinePaint(string xml, XmlRange outline)
    {
        var fills = DirectChildRanges(xml, outline).Where(entry => entry.LocalName is "noFill" or "solidFill").ToArray();
        return fills.Length == 1 && fills[0].LocalName == "solidFill" && TryNativeStyleXmlColor(xml, fills[0], "line", out _);
    }

    private static string? NativeStyleXmlAttribute(string xml, XmlRange range)
    {
        var fragment = xml[range.Start..range.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == range.LocalName);
        if (startTag is null) return null;
        var attributes = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == "val")
            .ToArray();
        return attributes.Length == 1 ? attributes[0].Groups["value"].Value : null;
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
            if (LeafKind(operation) is "chartTitleText" || IsChartDataLeafKind(LeafKind(operation)))
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
                        candidate.Binding.SeriesIndex == operation.ChartSeriesIndex &&
                        candidate.Binding.PointIndex == operation.ChartPointIndex &&
                        ChartDataChannel(LeafKind(operation)) == NativeChartDataChannel(candidate.Binding));
                    if (point is null || point.Binding.Value != operation.Value || point.Binding.Formula != operation.ChartFormula)
                        throw new CodecException("presentation_edit_verification_failed", $"PPTX chart data operation {operation.OperationId} did not survive package reopen in the ChartPart cache.", operation.TargetPartPath);
                    if (string.IsNullOrEmpty(operation.EmbeddedPackagePartPath))
                    {
                        if (chart.Binding.EmbeddedPackagePartPath.Length != 0 ||
                            chart.Binding.EmbeddedPackageSourceSha256.Length != 0 ||
                            chart.Binding.EmbeddedPackageRelationshipId.Length != 0 ||
                            point.Binding.WorksheetPartPath.Length != 0 ||
                            point.Binding.WorksheetSourceSha256.Length != 0 ||
                            point.Binding.WorksheetName.Length != 0 ||
                            point.Binding.CellReference.Length != 0)
                            throw new CodecException("presentation_edit_verification_failed", $"PPTX literal chart data operation {operation.OperationId} unexpectedly acquired an embedded-workbook binding.", operation.TargetPartPath);
                    }
                    else if (!chart.Binding.EmbeddedPackagePartPath.Equals(operation.EmbeddedPackagePartPath, StringComparison.OrdinalIgnoreCase) ||
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

    private static byte[] ReadEntry(ZipArchiveEntry entry)
    {
        using var input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] RequiredPart(IReadOnlyDictionary<string, byte[]> parts, string path) =>
        parts.TryGetValue(path, out var bytes)
            ? bytes
            : throw new CodecException("presentation_edit_target_missing", $"PPTX part {path} is missing.", path);

    private static string[] ChangedPartsStreaming(byte[] sourceBytes, byte[] outputBytes)
    {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourceArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read, leaveOpen: false);
        using var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Read, leaveOpen: false);
        var sourceEntries = sourceArchive.Entries
            .Where(entry => !entry.FullName.EndsWith("/", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        var outputEntries = outputArchive.Entries
            .Where(entry => !entry.FullName.EndsWith("/", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        return sourceEntries.Keys.Concat(outputEntries.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path =>
            {
                if (!sourceEntries.TryGetValue(path, out var sourceEntry) ||
                    !outputEntries.TryGetValue(path, out var outputEntry))
                    return true;
                using var left = sourceEntry.Open();
                using var right = outputEntry.Open();
                return !StreamsEqual(left, right);
            })
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string[] ChangedParts(byte[] sourceBytes, byte[] outputBytes) =>
        ChangedPartsStreaming(sourceBytes, outputBytes);

    private static bool ContainsPart(byte[] bytes, string path)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        return archive.GetEntry(path) is not null;
    }

    private static bool StreamsEqual(Stream left, Stream right)
    {
        Span<byte> leftBuffer = stackalloc byte[64 * 1024];
        Span<byte> rightBuffer = stackalloc byte[64 * 1024];
        while (true)
        {
            var leftRead = left.Read(leftBuffer);
            var rightRead = right.Read(rightBuffer);
            if (leftRead != rightRead) return false;
            if (leftRead == 0) return true;
            if (!leftBuffer[..leftRead].SequenceEqual(rightBuffer[..rightRead])) return false;
        }
    }

    private sealed class LazyPackageParts : IReadOnlyDictionary<string, byte[]>
    {
        private readonly byte[] _package;
        private readonly HashSet<string> _paths;
        private readonly Dictionary<string, byte[]> _cache = new(StringComparer.OrdinalIgnoreCase);

        internal LazyPackageParts(byte[] package)
        {
            _package = package;
            using var stream = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            _paths = archive.Entries
                .Where(entry => !entry.FullName.EndsWith("/", StringComparison.Ordinal))
                .Select(entry => entry.FullName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        public IEnumerable<string> Keys => _paths;
        public IEnumerable<byte[]> Values => _paths.Select(path => this[path]);
        public int Count => _paths.Count;
        public byte[] this[string key] =>
            TryGetValue(key, out var value)
                ? value
                : throw new KeyNotFoundException(key);

        public bool ContainsKey(string key) => _paths.Contains(key);

        public bool TryGetValue(string key, out byte[] value)
        {
            if (_cache.TryGetValue(key, out value!)) return true;
            if (!_paths.Contains(key))
            {
                value = [];
                return false;
            }
            using var stream = new MemoryStream(_package, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            var entry = archive.GetEntry(key);
            if (entry is null)
            {
                value = [];
                return false;
            }
            value = ReadEntry(entry);
            _cache[key] = value;
            return true;
        }

        public IEnumerator<KeyValuePair<string, byte[]>> GetEnumerator() =>
            _paths.Select(path => new KeyValuePair<string, byte[]>(path, this[path])).GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

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

    private static IReadOnlyList<XmlRange> ShapeElementRanges(
        string xml,
        string parentLocalName,
        bool includeGroupProperties = false)
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
        return includeGroupProperties
            ? children
            : children.Where(child => child.LocalName is not "nvGrpSpPr" and not "grpSpPr").ToArray();
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

    private static bool TryXmlAttributeValue(string xml, XmlRange range, string attribute, out string value)
    {
        value = string.Empty;
        var fragment = xml[range.Start..range.End];
        var startTag = XmlTokenPattern().Matches(fragment).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal) && LocalName(match.Value) == range.LocalName);
        if (startTag is null) return false;
        var matches = XmlAttributePattern().Matches(startTag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == attribute)
            .ToArray();
        if (matches.Length != 1) return false;
        value = matches[0].Groups["value"].Value;
        return true;
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
    private static bool ValidFontFamilyToken(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 255 || value.Trim() != value || value.StartsWith("+", StringComparison.Ordinal) || value.Any(char.IsControl)) return false;
        for (var index = 0; index < value.Length; index++)
        {
            var unit = value[index];
            if (char.IsHighSurrogate(unit))
            {
                if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1])) return false;
                index++;
            }
            else if (char.IsLowSurrogate(unit)) return false;
        }
        return true;
    }
    private static bool IsRgbToken(string value) => value.Length == 6 && value.All(Uri.IsHexDigit);
    private static bool ValidBooleanToken(string value) => value is "0" or "1";
    private static bool ValidOpacityToken(string value) =>
        uint.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var opacity) &&
        opacity <= 100_000 &&
        opacity.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    private static bool ValidGlowRadiusToken(string value) =>
        long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var radius) &&
        radius is >= 0 and <= 12_700_000 &&
        radius.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    private static bool ValidSoftEdgeRadiusToken(string value) =>
        long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var radius) &&
        radius is >= 0 and <= 12_700_000 &&
        radius.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    private static bool ValidInnerShadowGeometryToken(string kind, string value)
    {
        if (kind is "shapeInnerShadowBlurRadiusEmu" or "shapeInnerShadowDistanceEmu" or
            "imageInnerShadowBlurRadiusEmu" or "imageInnerShadowDistanceEmu")
        {
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var coordinate) &&
                coordinate is >= 0 and <= 2_147_483_647 &&
                coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        }
        return kind is "shapeInnerShadowDirectionDegrees" or "imageInnerShadowDirectionDegrees" &&
            long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
            direction is >= 0 and < 21_600_000 &&
            direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    }
    private static bool ValidReflectionGeometryToken(string kind, string value)
    {
        if (kind is "shapeReflectionBlurRadiusEmu" or "imageReflectionBlurRadiusEmu")
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var blur) &&
                blur is >= 0 and <= 12_700_000 &&
                blur.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        if (kind is "shapeReflectionDistanceEmu" or "imageReflectionDistanceEmu")
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var distance) &&
                distance is >= 0 and <= 1_270_000_000 &&
                distance.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        return kind is "shapeReflectionDirectionDegrees" or "imageReflectionDirectionDegrees" &&
            long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
            direction is >= 0 and < 21_600_000 &&
            direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    }
    private static bool ValidTextInnerShadowGeometryToken(string kind, string value)
    {
        if (kind is "textInnerShadowBlurRadiusEmu" or "textDefaultInnerShadowBlurRadiusEmu" or "textInnerShadowDistanceEmu" or "textDefaultInnerShadowDistanceEmu")
        {
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var coordinate) &&
                coordinate is >= 0 and <= 2_147_483_647 &&
                coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        }
        return kind is ("textInnerShadowDirectionDegrees" or "textDefaultInnerShadowDirectionDegrees") &&
            long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
            direction is >= 0 and < 21_600_000 &&
            direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    }
    private static bool ValidTextReflectionGeometryToken(string kind, string value)
    {
        if (kind is "textReflectionBlurRadiusEmu" or "textDefaultReflectionBlurRadiusEmu" or "textDefaultReflectionDistanceEmu" or "textReflectionDistanceEmu")
        {
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var coordinate) &&
                coordinate is >= 0 and <= 2_147_483_647 &&
                coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        }
        return kind is ("textReflectionDirectionDegrees" or "textDefaultReflectionDirectionDegrees") &&
            long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
            direction is >= 0 and < 21_600_000 &&
            direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    }
    private static bool ValidTextSoftEdgeRadiusToken(string value) =>
        long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var radius) &&
        radius is >= 0 and <= 12_700_000 &&
        radius.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
    private static bool ValidShadowGeometryToken(string kind, string value)
    {
        if (kind is "shadowBlurRadiusEmu" or "imageShadowBlurRadiusEmu" or "shadowDistanceEmu" or "imageShadowDistanceEmu" or "textDefaultShadowBlurRadiusEmu" or "textDefaultShadowDistanceEmu")
        {
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var coordinate) &&
                coordinate is >= 0 and <= 2_147_483_647 &&
                coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        }
        if (kind is "shadowDirectionDegrees" or "imageShadowDirectionDegrees" or "textDefaultShadowDirectionDegrees")
        {
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
                direction is >= 0 and < 21_600_000 &&
                direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        }
        return (kind is "shadowAlignment" or "imageShadowAlignment" or "textDefaultShadowAlignment") && (value is "tl" or "t" or "tr" or "l" or "ctr" or "r" or "bl" or "b" or "br");
    }
    private static bool TryCanonicalBoolean(string value, out string canonical)
    {
        switch (System.Net.WebUtility.HtmlDecode(value).ToLowerInvariant())
        {
            case "0":
            case "false":
                canonical = "0";
                return true;
            case "1":
            case "true":
                canonical = "1";
                return true;
            default:
                canonical = string.Empty;
                return false;
        }
    }
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
    private static bool IsChartDataLeafKind(string leafKind) => PpjNativeLeafProjection.IsChartDataLeafKind(leafKind);
    private static string ChartDataChannel(string leafKind) => leafKind switch
    {
        "chartDataCategory" => "category",
        "chartDataXValue" => "x",
        "chartDataYValue" => "y",
        "chartDataBubbleSize" => "size",
        _ => "value",
    };
    private static string NativeChartDataChannel(PresentationNativeChartDataPoint point) =>
        string.IsNullOrEmpty(point.Channel) ? "value" : point.Channel;
    private static string MutationPartPath(PresentationEditOperation operation) =>
        LeafKind(operation) is "chartTitleText" or "diagramText" || IsChartDataLeafKind(LeafKind(operation)) ? operation.TargetPartPath : operation.SlidePartPath;
    private static bool IsGeometryLeaf(string leafKind) =>
        leafKind is "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical";
    private static bool IsGroupChildGeometryLeaf(string leafKind) =>
        leafKind is "childLeftEmu" or "childTopEmu" or "childWidthEmu" or "childHeightEmu";
    private static bool LeafValuesEqual(string left, string right, string leafKind) =>
        leafKind is "fillRgb" or "lineRgb" or "shape3dContourRgb" or "shape3dExtrusionRgb"
            ? PptxColor.Normalize(left) == PptxColor.Normalize(right)
            : leafKind is "lineScheme" or "fillScheme" or "shape3dContourColorScheme" or "shape3dExtrusionColorScheme"
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
