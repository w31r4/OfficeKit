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
            if (leafKind is not ("text" or "tableCellText" or "nativeText" or "paragraphAlignment" or "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" or "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" or "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier" or "paragraphMarginLeftEmu" or "paragraphIndentEmu" or "paragraphBulletCharacter" or "paragraphBulletAutoNumberScheme" or "paragraphBulletAutoNumberStartAt" or "paragraphBulletFontFamily" or "paragraphBulletColorRgb" or "paragraphBulletColorScheme" or "paragraphBulletSizePoints" or "paragraphBulletSizePercent" or "paragraphLevel" or "verticalAnchor" or "textBodyInsetLeftEmu" or "textBodyInsetTopEmu" or "textBodyInsetRightEmu" or "textBodyInsetBottomEmu" or "textBodyWrap" or "textBodyColumnCount" or "textBodyAutoFit" or "textBodyColumnDirection" or "textBodyVerticalText" or "fontSizePoints" or "fontFamily" or "fontFamilyEastAsia" or "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontColorRgb" or "fontColorScheme" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps" or "fontHighlightRgb" or "fontHighlightScheme" or "fillRgb" or "fillOpacityThousandthPercent" or "shadowOpacityThousandthPercent" or "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment" or "shadowColorRgb" or "shadowColorScheme" or "imageOpacityThousandthPercent" or "imageMaskPreset" or "fillScheme" or "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu" or "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical" or "imageAsset" or "imageSvgAsset" or "chartTitleText" or "chartDataValue" or "diagramText" or "deleteElement"))
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
            if (leafKind == "fontSizePoints")
            {
                if (!uint.TryParse(operation.ExpectedValue, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var expectedFontSize) ||
                    !uint.TryParse(operation.Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var requestedFontSize) ||
                    expectedFontSize == 0 || expectedFontSize > 76_800 || requestedFontSize == 0 || requestedFontSize > 76_800)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} font size must be an integer from 1 through 76800 hundredths of a point.");
            }
            if (leafKind is "fontFamily" or "fontFamilyEastAsia")
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
            if (leafKind == "textBodyAutoFit")
            {
                var expected = ParseTextBodyAutoFitToken(operation.ExpectedValue, operation);
                var requested = ParseTextBodyAutoFitToken(operation.Value, operation);
                if (expected == requested)
                    throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its text-body AutoFit mode.");
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
            if (leafKind is "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment")
            {
                if (!ValidShadowGeometryToken(leafKind, operation.ExpectedValue) ||
                    !ValidShadowGeometryToken(leafKind, operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} {leafKind} must use a changed bounded native shadow geometry token.");
            }
            if (leafKind == "shadowOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} shadowOpacityThousandthPercent must use a changed canonical value from 0 through 100000.");
            }
            if (leafKind is "shadowColorRgb" or "shadowColorScheme")
            {
                if (leafKind == "shadowColorRgb")
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
            if (leafKind == "imageOpacityThousandthPercent")
            {
                if (!ValidOpacityToken(operation.ExpectedValue) || !ValidOpacityToken(operation.Value) ||
                    operation.ExpectedValue == operation.Value)
                    throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} imageOpacityThousandthPercent must use a changed canonical value from 0 through 100000.");
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
                 ((projectedElement.Source.Editable &&
                   (LeafKind(operation) is "fillRgb" or "fillScheme" or "fillOpacityThousandthPercent" or "shadowOpacityThousandthPercent" or "shadowColorRgb" or "shadowColorScheme" or "lineRgb" or "lineScheme" or "lineWidthEmu" or "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical" ||
                    (LeafKind(operation) is "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow") && HasSafeNativeShapeStyle(shape, LeafKind(operation)) ||
                    (LeafKind(operation) is "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment") && shape.ShapeProperties is not null && HasSafeNativeShadowGeometry(shape.ShapeProperties, LeafKind(operation)))) ||
                 (!projectedElement.Source.Editable && LeafKind(operation) is ("fillRgb" or "fillOpacityThousandthPercent" or "shadowOpacityThousandthPercent" or "shadowColorRgb" or "shadowColorScheme" or "fillScheme" or "lineRgb" or "lineScheme" or "lineWidthEmu" or "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment") && HasSafeNativeShapeStyle(shape, LeafKind(operation))) ||
                 (projectedElement.Source.TextEditable && LeafKind(operation) is ("text" or "paragraphAlignment" or "paragraphLineSpacingPoints" or "paragraphLineSpacingMultiplier" or "paragraphSpaceBeforePoints" or "paragraphSpaceBeforeMultiplier" or "paragraphSpaceAfterPoints" or "paragraphSpaceAfterMultiplier" or "paragraphMarginLeftEmu" or "paragraphIndentEmu" or "paragraphBulletCharacter" or "paragraphBulletAutoNumberScheme" or "paragraphBulletAutoNumberStartAt" or "paragraphBulletFontFamily" or "paragraphBulletColorRgb" or "paragraphBulletColorScheme" or "paragraphBulletSizePoints" or "paragraphBulletSizePercent" or "paragraphLevel" or "verticalAnchor" or "textBodyInsetLeftEmu" or "textBodyInsetTopEmu" or "textBodyInsetRightEmu" or "textBodyInsetBottomEmu" or "textBodyWrap" or "textBodyColumnCount" or "textBodyAutoFit" or "textBodyColumnDirection" or "textBodyVerticalText" or "fontSizePoints" or "fontFamily" or "fontFamilyEastAsia" or "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontColorRgb" or "fontColorScheme" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps" or "fontHighlightRgb" or "fontHighlightScheme" or "rotationDegrees" or "flipHorizontal" or "flipVertical") && PptxCodec.SupportsBoundTextLeaf(shape))))
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
                     (LeafKind(operation) is "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu") &&
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
                      LeafKind(operation) == "imageMaskPreset" && HasSafeNativePictureMask(picture)))
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
            if (leafKind == "textBodyAutoFit")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} textBodyAutoFit target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(CompileTextBodyAutoFitXmlPatch(xml, range, proof));
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
            if ((leafKind is "rotationDegrees" or "flipHorizontal" or "flipVertical") && range.LocalName is not ("sp" or "pic"))
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
            if (leafKind is "fontSizePoints" or "fontFamily" or "fontFamilyEastAsia" or "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontColorRgb" or "fontColorScheme" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps" or "fontHighlightRgb" or "fontHighlightScheme")
            {
                if (range.LocalName != "sp")
                    throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {leafKind} target has the wrong native element type.", operation.SlidePartPath);
                patches.Add(leafKind == "fontSizePoints"
                    ? CompileTextFontSizeXmlPatch(xml, range, proof, drawingPrefixes)
                    : leafKind is "fontFamily" or "fontFamilyEastAsia"
                        ? CompileTextFontFamilyXmlPatch(xml, range, proof, drawingPrefixes)
                        : leafKind is "fontLanguage" or "fontBold" or "fontItalic" or "fontUnderline" or "fontStrike" or "fontKerningPoints" or "fontBaselinePercent" or "fontSpacingPoints" or "fontCaps"
                            ? CompileTextFontBooleanXmlPatch(xml, range, proof, drawingPrefixes)
                            : CompileTextFontColorXmlPatch(xml, range, proof, drawingPrefixes));
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
    }

    private static bool HasSafeNativeConnectorLine(P.ConnectionShape connector, string kind)
    {
        var outline = connector.ShapeProperties?.Elements<A.Outline>().ToArray();
        if (outline is not { Length: 1 }) return false;
        if (kind == "lineWidthEmu")
            return outline[0].Width?.Value is { } width && width is >= 0 and <= 20_116_800;
        if (kind is "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow")
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
            return hasExplicitPaint && PptxLineStyleCodec.TryReadJoinLeaf(outline[0], out _);
        }
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
        if (kind is "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment")
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
        if (kind is not ("lineRgb" or "lineScheme")) return false;
        return linePaints.Length == 1 && linePaints[0] is A.SolidFill directLineSolid &&
            (kind == "lineRgb" && HasSafeNativeRgbFill(directLineSolid) ||
             kind == "lineScheme" && HasSafeNativeSchemeFill(directLineSolid));
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
            "shadowBlurRadiusEmu" => "blurRad",
            "shadowDistanceEmu" => "dist",
            "shadowDirectionDegrees" => "dir",
            "shadowAlignment" => "algn",
            _ => string.Empty,
        };
        var attributes = outer.GetAttributes().Where(attribute => attribute.LocalName == attributeName).ToArray();
        if (attributes.Length != 1 || string.IsNullOrEmpty(attributes[0].Value)) return false;
        if (kind is "shadowBlurRadiusEmu" or "shadowDistanceEmu")
            return long.TryParse(attributes[0].Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var coordinate) &&
                coordinate is >= 0 and <= 2_147_483_647 &&
                coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) == attributes[0].Value;
        if (kind == "shadowDirectionDegrees")
            return long.TryParse(attributes[0].Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
                direction is >= 0 and < 21_600_000 &&
                direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == attributes[0].Value;
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
        var colorMatches = kind == "shadowColorRgb"
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
        if (kind is "fontFamily" or "fontFamilyEastAsia")
        {
            if (element is not P.Shape shape)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} {kind} target is not a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length || leaves[operation.TextLeafIndex].Parent is not A.Run run)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit run font family.", operation.SlidePartPath);
            var runProperties = run.RunProperties;
            OpenXmlElement[] fonts = kind == "fontFamily"
                ? runProperties?.Elements<A.LatinFont>().Cast<OpenXmlElement>().ToArray() ?? []
                : runProperties?.Elements<A.EastAsianFont>().Cast<OpenXmlElement>().ToArray() ?? [];
            var typeface = kind == "fontFamily"
                ? (fonts.FirstOrDefault() as A.LatinFont)?.Typeface?.Value
                : (fonts.FirstOrDefault() as A.EastAsianFont)?.Typeface?.Value;
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
        if (kind is "shadowColorRgb" or "shadowColorScheme")
        {
            if (element is not P.Shape shape || shape.ShapeProperties is null || !HasSafeNativeShadowColor(shape.ShapeProperties, kind))
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit shadow color.", operation.SlidePartPath);
            if (!PptxShadowCodec.TryRead(shape.ShapeProperties, out var shadow) || shadow is null)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no bounded explicit shadow color.", operation.SlidePartPath);
            var value = kind == "shadowColorRgb" ? shadow.ColorRgb : shadow.ColorScheme;
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
        if (kind is "shadowBlurRadiusEmu" or "shadowDistanceEmu" or "shadowDirectionDegrees" or "shadowAlignment")
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
                _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported shadow geometry.", operation.SlidePartPath),
            };
            var attribute = outerShadow.GetAttributes().Single(item => item.LocalName == attributeName);
            return attribute.Value ?? throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} shadow geometry attribute has no value.", operation.SlidePartPath);
        }
        if ((kind is "fillRgb" or "fillScheme" or "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu") && element is P.GroupShape group)
        {
            if (!PptxNativeStyleLeafCodec.TryResolve(group, operation.NativeLeafIndex, out var leaf) || leaf.Kind != kind)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} target is not a bounded source-bound group style leaf.", operation.SlidePartPath);
            return leaf.Value;
        }
        var properties = element switch
        {
            P.Shape shape => shape.ShapeProperties,
            P.ConnectionShape connector when kind is "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu" => connector.ShapeProperties,
            P.Picture picture when IsGeometryLeaf(kind) => picture.ShapeProperties,
            _ => null,
        } ??
            throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} target has no p:spPr.", operation.SlidePartPath);
        var transform = properties.Transform2D;
        return kind switch
        {
            "fillRgb" => RequiredLeafValue(PptxColor.SolidRgb(properties.GetFirstChild<A.SolidFill>()), operation),
            "fillOpacityThousandthPercent" => ReadFillOpacity(properties.GetFirstChild<A.SolidFill>(), operation),
            "fillScheme" => RequiredLeafValue(PptxColor.SolidSchemeWithOpacity(properties.GetFirstChild<A.SolidFill>()), operation),
            "lineRgb" => RequiredLeafValue(PptxColor.SolidRgb(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()), operation),
            "lineScheme" => RequiredLeafValue(NativeSchemeToken(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(), operation), operation),
            "lineStyle" => RequiredLeafValue(ReadLineStyle(properties.GetFirstChild<A.Outline>(), operation), operation),
            "lineCap" => RequiredLeafValue(ReadLineCap(properties.GetFirstChild<A.Outline>(), operation), operation),
            "lineJoin" => RequiredLeafValue(ReadLineJoin(properties.GetFirstChild<A.Outline>(), operation), operation),
            "lineStartArrow" => RequiredLeafValue(ReadLineArrow(properties.GetFirstChild<A.Outline>(), operation, start: true), operation),
            "lineEndArrow" => RequiredLeafValue(ReadLineArrow(properties.GetFirstChild<A.Outline>(), operation, start: false), operation),
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

    private static string ParagraphAlignmentName(A.TextAlignmentTypeValues value) =>
        value == A.TextAlignmentTypeValues.Left ? "left" :
        value == A.TextAlignmentTypeValues.Center ? "center" :
        value == A.TextAlignmentTypeValues.Right ? "right" :
        value == A.TextAlignmentTypeValues.Justified ? "justify" : string.Empty;

    private static string ParagraphAlignmentName(string value) =>
        value switch
        {
            "l" => "left",
            "ctr" => "center",
            "r" => "right",
            "just" => "justify",
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
        if (owner == "grpSp" && (LeafKind(operation) is "fillRgb" or "fillScheme" or "lineRgb" or "lineScheme" or "lineStyle" or "lineCap" or "lineJoin" or "lineStartArrow" or "lineEndArrow" or "lineWidthEmu"))
            return CompileNativeStyleXmlPatch(xml, elementRange, proof);
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
            case "shadowBlurRadiusEmu":
            case "shadowDistanceEmu":
            case "shadowDirectionDegrees":
            case "shadowAlignment":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                var shadowEffectList = DirectChildRange(xml, properties, "spPr", "effectLst", operation);
                leaf = DirectChildRange(xml, shadowEffectList, "effectLst", "outerShdw", operation);
                attribute = LeafKind(operation) switch
                {
                    "shadowBlurRadiusEmu" => "blurRad",
                    "shadowDistanceEmu" => "dist",
                    "shadowDirectionDegrees" => "dir",
                    "shadowAlignment" => "algn",
                    _ => throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has unsupported shadow geometry.", operation.SlidePartPath),
                };
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
            case "fillScheme":
                if (owner != "sp") throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose fillScheme.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, DirectChildRange(xml, properties, "spPr", "solidFill", operation), "solidFill", "schemeClr", operation);
                attribute = "val";
                break;
            case "lineRgb":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineRgb.", operation.SlidePartPath);
                var outline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                leaf = DirectChildRange(xml, DirectChildRange(xml, outline, "ln", "solidFill", operation), "solidFill", "srgbClr", operation);
                attribute = "val";
                break;
            case "lineScheme":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineScheme.", operation.SlidePartPath);
                var schemeOutline = DirectChildRange(xml, properties, "spPr", "ln", operation);
                leaf = DirectChildRange(xml, DirectChildRange(xml, schemeOutline, "ln", "solidFill", operation), "solidFill", "schemeClr", operation);
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
            case "lineWidthEmu":
                if (owner is not ("sp" or "cxnSp")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose lineWidthEmu.", operation.SlidePartPath);
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
            case "rotationDegrees":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose rotationDegrees.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "xfrm", operation);
                attribute = "rot";
                break;
            case "flipHorizontal":
            case "flipVertical":
                if (owner is not ("sp" or "pic")) throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} target does not expose {LeafKind(operation)}.", operation.SlidePartPath);
                leaf = DirectChildRange(xml, properties, "spPr", "xfrm", operation);
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
            _ => operation.ExpectedValue,
        };
        var replacement = LeafKind(operation) switch
        {
            "lineStyle" => PptxLineStyleCodec.TryPresetDashToken(operation.Value, out var requestedStyleToken) ? requestedStyleToken : string.Empty,
            "lineCap" => PptxLineStyleCodec.TryCapToken(operation.Value, out var requestedCapToken) ? requestedCapToken : string.Empty,
            "imageMaskPreset" => PptxCustomGeometryCodec.TryPreset(operation.Value, out _) ? operation.Value : string.Empty,
            _ => operation.Value,
        };
        var matches = LeafKind(operation) is "flipHorizontal" or "flipVertical"
            ? TryCanonicalBoolean(valueGroup.Value, out var actualBoolean) && actualBoolean == operation.ExpectedValue
            : valueGroup.Value == expectedScalar;
        if (!matches)
            throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw scalar does not match the expected value.", operation.SlidePartPath);
        var start = leaf.Start + startTag.Index + valueGroup.Index;
        return new PptxXmlPatch(operation, start, start + valueGroup.Length, replacement, proof.SourceElementSha256, proof.MutationPartPath);
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
        var fontName = LeafKind(operation) == "fontFamily" ? "latin" : "ea";
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
        var children = ShapeElementRanges(fragment, parentLocalName).Where(child => child.LocalName == childLocalName).ToArray();
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
    private static bool ValidShadowGeometryToken(string kind, string value)
    {
        if (kind is "shadowBlurRadiusEmu" or "shadowDistanceEmu")
        {
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var coordinate) &&
                coordinate is >= 0 and <= 2_147_483_647 &&
                coordinate.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        }
        if (kind == "shadowDirectionDegrees")
        {
            return long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var direction) &&
                direction is >= 0 and < 21_600_000 &&
                direction.ToString(System.Globalization.CultureInfo.InvariantCulture) == value;
        }
        return kind == "shadowAlignment" && (value is "tl" or "t" or "tr" or "l" or "ctr" or "r" or "bl" or "b" or "br");
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
    private static string MutationPartPath(PresentationEditOperation operation) =>
        LeafKind(operation) is "chartTitleText" or "chartDataValue" or "diagramText" ? operation.TargetPartPath : operation.SlidePartPath;
    private static bool IsGeometryLeaf(string leafKind) =>
        leafKind is "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical";
    private static bool LeafValuesEqual(string left, string right, string leafKind) =>
        leafKind is "fillRgb" or "lineRgb"
            ? PptxColor.Normalize(left) == PptxColor.Normalize(right)
            : leafKind is "lineScheme" or "fillScheme"
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
