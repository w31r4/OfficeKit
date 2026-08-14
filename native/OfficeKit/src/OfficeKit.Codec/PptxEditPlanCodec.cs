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
    string SourceElementSha256);

internal sealed record PptxXmlPatch(
    PresentationEditOperation Operation,
    int Start,
    int End,
    string Replacement,
    string SourceElementSha256);

// Applies a finite, source-bound edit plan directly to the original XML token
// stream. The Open XML SDK is used only as an independent structural oracle;
// it never serializes the mutated SlidePart.
internal static partial class PptxEditPlanCodec
{
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    [GeneratedRegex("^ppt/slides/slide[1-9][0-9]*[.]xml$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex SlidePathPattern();

    [GeneratedRegex("<!--.*?-->|<!\\[CDATA\\[.*?\\]\\]>|<\\?.*?\\?>|</?(?:[A-Za-z_][\\w.-]*:)?[A-Za-z_][\\w.-]*\\b[^>]*>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex XmlTokenPattern();

    [GeneratedRegex("^</?(?:[A-Za-z_][\\w.-]*:)?(?<name>[A-Za-z_][\\w.-]*)\\b", RegexOptions.CultureInvariant)]
    private static partial Regex XmlLocalNamePattern();

    [GeneratedRegex("xmlns:(?<prefix>[A-Za-z_][\\w.-]*)\\s*=\\s*(?<quote>['\"])(?<uri>.*?)\\k<quote>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex NamespacePattern();

    [GeneratedRegex("(?<open><(?<prefix>[A-Za-z_][\\w.-]*):t\\b[^>]*>)(?<value>.*?)(?<close></\\k<prefix>:t\\s*>)", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex TextLeafPattern();

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
        var proofs = ProveOperations(sourceBytes, request, sourceProjection);
        var sourceParts = PackageParts(sourceBytes);
        var patchedParts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var results = new List<PresentationEditOperationResult>();

        foreach (var group in proofs.GroupBy(proof => proof.Operation.SlidePartPath, StringComparer.OrdinalIgnoreCase))
        {
            var partPath = group.Key;
            var sourcePart = sourceParts[partPath];
            var patches = CompileXmlPatches(sourcePart, group.ToArray());
            var outputPart = ApplyPatches(sourcePart, patches, results);
            if (sourcePart.AsSpan().SequenceEqual(outputPart))
                throw new CodecException("presentation_edit_plan_noop", $"PPTX edit plan produced no change for {partPath}.", partPath);
            patchedParts.Add(partPath, outputPart);
        }

        var outputBytes = ReplaceParts(sourceBytes, patchedParts);
        var changedParts = ChangedParts(sourceBytes, outputBytes);
        var expectedParts = patchedParts.Keys.OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
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

        _ = PackageGuards.ValidateAndCollectOpaque(outputBytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        var sourceValidationWarnings = PptxCodec.ValidateEditPlanOutput(sourceBytes, outputBytes, limits);
        VerifyOutput(outputBytes, request, results);

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
        if (request.Operations.GroupBy(operation => (operation.SlidePartPath.ToLowerInvariant(), ShapeTreePathKey(operation), operation.TextLeafIndex)).Any(group => group.Count() > 1))
            throw new CodecException("duplicate_presentation_edit_target", "PPTX edit plan cannot edit the same text leaf twice.");
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
            if (!IsSha256(operation.ExpectedSlideSha256) || !IsSha256(operation.ExpectedElementSha256) ||
                !IsSha256(operation.ExpectedSemanticSha256) || !IsSha256(operation.ExpectedTextSha256))
                throw new CodecException("invalid_presentation_edit_precondition", $"PPTX edit operation {operation.OperationId} requires SHA-256 preconditions.");
            if (!Hash(Encoding.UTF8.GetBytes(operation.ExpectedValue)).Equals(operation.ExpectedTextSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_text_hash_mismatch", $"PPTX edit operation {operation.OperationId} expected text does not match expected_text_sha256.");
            if (operation.ExpectedValue == operation.Value)
                throw new CodecException("presentation_edit_plan_noop", $"PPTX edit operation {operation.OperationId} must change its target value.");
            textBudget = checked(textBudget + (ulong)operation.ExpectedValue.Length + (ulong)operation.Value.Length);
            if (textBudget > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"PPTX edit-plan text exceeds max_cells ({limits.MaxCells}).");
        }
    }

    private static IReadOnlyList<PptxEditPlanProof> ProveOperations(
        byte[] sourceBytes,
        PresentationEditPlanRequest request,
        PresentationArtifact sourceProjection)
    {
        using var stream = new MemoryStream(sourceBytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false, new OpenSettings { AutoSave = false });
        var presentationPart = package.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.");
        var slideByPath = presentationPart.SlideParts.ToDictionary(PartPath, StringComparer.OrdinalIgnoreCase);
        var projectedSlideByPath = sourceProjection.Slides.ToDictionary(slide => slide.Source.PartPath, StringComparer.OrdinalIgnoreCase);
        var proofs = new List<PptxEditPlanProof>();
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
            if (element is not P.Shape shape || !PptxCodec.SupportsBoundTextLeaf(shape))
                throw new CodecException("unsupported_presentation_edit", $"PPTX edit operation {operation.OperationId} target is not a safely editable text shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (operation.TextLeafIndex >= (uint)leaves.Length)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} text-leaf index is out of range.", operation.SlidePartPath);
            if (leaves[operation.TextLeafIndex].Text != operation.ExpectedValue)
                throw new CodecException("presentation_text_precondition_failed", $"PPTX edit operation {operation.OperationId} old text does not match the source leaf.", operation.SlidePartPath);
            proofs.Add(new PptxEditPlanProof(operation, elementHash));
        }
        return proofs;
    }

    private static PptxXmlPatch[] CompileXmlPatches(byte[] partBytes, IReadOnlyList<PptxEditPlanProof> proofs)
    {
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
            if (range.LocalName != "sp")
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} raw target is not p:sp.", operation.SlidePartPath);
            var elementXml = xml[range.Start..range.End];
            var leaves = TextLeafPattern().Matches(elementXml)
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
                proof.SourceElementSha256));
        }
        var ordered = patches.OrderBy(patch => patch.Start).ToArray();
        for (var index = 1; index < ordered.Length; index++)
            if (ordered[index - 1].End > ordered[index].Start)
                throw new CodecException("overlapping_presentation_edit_operations", "PPTX edit plan operations overlap in the source XML.", ordered[index].Operation.SlidePartPath);
        return ordered;
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
            var outputEndCharacter = patch.Start + patch.Replacement.Length;
            var outputEnd = checked((ulong)(bomBytes + StrictUtf8.GetByteCount(output[..outputEndCharacter])));
            var result = new PresentationEditOperationResult
            {
                OperationId = patch.Operation.OperationId,
                SlideId = patch.Operation.SlideId,
                SlidePartPath = patch.Operation.SlidePartPath,
                TargetId = patch.Operation.TargetId,
                ShapeTreeIndex = patch.Operation.ShapeTreeIndex,
                TextLeafIndex = patch.Operation.TextLeafIndex,
                SourceElementSha256 = patch.SourceElementSha256,
                OldValueSha256 = Hash(Encoding.UTF8.GetBytes(patch.Operation.ExpectedValue)),
                NewValueSha256 = Hash(Encoding.UTF8.GetBytes(patch.Operation.Value)),
                SourceStartOffset = sourceStart,
                SourceEndOffset = sourceEnd,
                OutputEndOffset = outputEnd,
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
        IReadOnlyList<PresentationEditOperationResult> results)
    {
        using var stream = new MemoryStream(outputBytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false, new OpenSettings { AutoSave = false });
        var slideByPath = package.PresentationPart!.SlideParts.ToDictionary(PartPath, StringComparer.OrdinalIgnoreCase);
        var resultById = results.ToDictionary(result => result.OperationId, StringComparer.Ordinal);
        foreach (var operation in request.Operations)
        {
            var tree = slideByPath[operation.SlidePartPath].Slide!.CommonSlideData!.ShapeTree!;
            var element = ResolveShapeTreeElement(tree, ShapeTreePath(operation), operation);
            var shape = element as P.Shape ?? throw new CodecException("presentation_edit_verification_failed", "PPTX edited target is no longer a shape.", operation.SlidePartPath);
            var leaves = shape.Descendants<A.Text>().ToArray();
            if (leaves[operation.TextLeafIndex].Text != operation.Value)
                throw new CodecException("presentation_edit_verification_failed", $"PPTX edit operation {operation.OperationId} did not survive package reopen.", operation.SlidePartPath);
            resultById[operation.OperationId].OutputElementSha256 = HashElement(element);
        }
    }

    private static byte[] ReplaceParts(byte[] sourceBytes, IReadOnlyDictionary<string, byte[]> replacements)
    {
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
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

    private static byte[] ReadPart(byte[] bytes, string path)
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

    private static string[] ChangedParts(byte[] sourceBytes, byte[] outputBytes)
    {
        var source = PackageParts(sourceBytes).ToDictionary(entry => entry.Key, entry => Hash(entry.Value), StringComparer.OrdinalIgnoreCase);
        var output = PackageParts(outputBytes).ToDictionary(entry => entry.Key, entry => Hash(entry.Value), StringComparer.OrdinalIgnoreCase);
        return source.Keys.Concat(output.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path => !source.TryGetValue(path, out var left) || !output.TryGetValue(path, out var right) || left != right)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
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
