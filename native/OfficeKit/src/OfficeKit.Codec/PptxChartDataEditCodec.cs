using System.Text.RegularExpressions;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static partial class PptxEditPlanCodec
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [GeneratedRegex("^</?(?:(?<prefix>[A-Za-z_][\\w.-]*):)?(?<name>[A-Za-z_][\\w.-]*)\\b", RegexOptions.CultureInvariant)]
    private static partial Regex QualifiedTagPattern();

    [GeneratedRegex("\\bxmlns\\s*=\\s*(?<quote>['\"])(?<uri>.*?)\\k<quote>", RegexOptions.CultureInvariant | RegexOptions.Singleline)]
    private static partial Regex DefaultNamespacePattern();

    [GeneratedRegex("^-?(?:0|[1-9][0-9]*)(?:[.][0-9]+)?(?:[Ee][+-]?[0-9]+)?$", RegexOptions.CultureInvariant)]
    private static partial Regex FiniteNumberPattern();

    private sealed record QualifiedOpenElement(string Prefix, string Name, int Start);

    private static PptxXmlPatch[] CompileChartDataXmlPatches(
        byte[] partBytes,
        IReadOnlyList<PptxEditPlanProof> proofs)
    {
        var (xml, _) = DecodeXml(partBytes);
        var chartPrefixes = NamespacePrefixes(xml, ChartNamespace);
        if (chartPrefixes.Count == 0)
            throw new CodecException("presentation_edit_target_missing", "PPTX ChartPart does not declare the chart namespace.", proofs[0].MutationPartPath);
        var series = NamespaceElementRanges(xml, "ser", chartPrefixes);
        var allRanges = NamespaceElementRanges(xml, null, chartPrefixes);
        var patches = new List<PptxXmlPatch>();
        foreach (var proof in proofs)
        {
            var operation = proof.Operation;
            if (operation.ChartSeriesIndex >= (uint)series.Count)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} chart series index is out of range.", proof.MutationPartPath);
            var seriesRange = series[(int)operation.ChartSeriesIndex];
            var numberReferences = allRanges.Where(range => range.LocalName == "numRef" && Contains(seriesRange, range)).ToArray();
            var matchingReferences = numberReferences.Where(reference =>
            {
                var formulas = allRanges.Where(range => range.LocalName == "f" && Contains(reference, range)).ToArray();
                return formulas.Length == 1 && ElementText(xml, formulas[0]) == EscapeText(operation.ChartFormula);
            }).ToArray();
            if (matchingReferences.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} chart formula is missing or ambiguous.", proof.MutationPartPath);
            var numberReference = matchingReferences[0];
            var caches = allRanges.Where(range => range.LocalName == "numCache" && Contains(numberReference, range)).ToArray();
            if (caches.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} requires one numeric cache.", proof.MutationPartPath);
            var points = allRanges.Where(range => range.LocalName == "pt" && Contains(caches[0], range) &&
                AttributeValue(xml, range, "idx") == operation.ChartPointIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            if (points.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} chart cache point is missing or ambiguous.", proof.MutationPartPath);
            var values = allRanges.Where(range => range.LocalName == "v" && Contains(points[0], range)).ToArray();
            if (values.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} chart cache value is missing or ambiguous.", proof.MutationPartPath);
            var (start, end) = ElementTextSpan(xml, values[0]);
            if (xml[start..end] != EscapeText(operation.ExpectedValue))
                throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} chart cache old value does not match.", proof.MutationPartPath);
            patches.Add(new PptxXmlPatch(operation, start, end, EscapeText(operation.Value), proof.SourceElementSha256, proof.MutationPartPath));
        }
        return OrderedNonOverlapping(patches, proofs[0].MutationPartPath);
    }

    private static void ApplyEmbeddedWorkbookPatches(
        IReadOnlyDictionary<string, byte[]> sourceParts,
        IReadOnlyList<PptxEditPlanProof> proofs,
        IDictionary<string, byte[]> patchedParts,
        IReadOnlyList<PresentationEditOperationResult> results)
    {
        var dataProofs = proofs.Where(proof => LeafKind(proof.Operation) == "chartDataValue").ToArray();
        foreach (var packageGroup in dataProofs.GroupBy(proof => proof.Operation.EmbeddedPackagePartPath, StringComparer.OrdinalIgnoreCase))
        {
            var packagePath = packageGroup.Key;
            if (!sourceParts.TryGetValue(packagePath, out var sourcePackage) ||
                !Hash(sourcePackage).Equals(packageGroup.First().Operation.ExpectedEmbeddedPackageSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_chart_data_binding_mismatch", "PPTX embedded workbook bytes changed after planning.", packagePath);
            var innerParts = PackageParts(sourcePackage);
            var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var worksheetGroup in packageGroup.GroupBy(proof => proof.Operation.EmbeddedWorksheetPartPath, StringComparer.OrdinalIgnoreCase))
            {
                var worksheetPath = worksheetGroup.Key;
                if (!innerParts.TryGetValue(worksheetPath, out var sourceWorksheet) ||
                    !Hash(sourceWorksheet).Equals(worksheetGroup.First().Operation.ExpectedEmbeddedWorksheetSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_chart_data_binding_mismatch", "PPTX embedded worksheet bytes changed after planning.", worksheetPath);
                var worksheetProofs = worksheetGroup.Select(proof => proof with { MutationPartPath = worksheetPath }).ToArray();
                var patches = CompileWorksheetValueXmlPatches(sourceWorksheet, worksheetProofs);
                var nestedResults = new List<PresentationEditOperationResult>();
                replacements.Add(worksheetPath, ApplyPatches(sourceWorksheet, patches, nestedResults));
                foreach (var nestedResult in nestedResults)
                {
                    var primary = results.Single(result => result.OperationId == nestedResult.OperationId);
                    primary.NestedFootprints.Add(new PresentationNestedMutationFootprint
                    {
                        ContainerPartPath = packagePath,
                        PartPath = worksheetPath,
                        OldValueSha256 = nestedResult.OldValueSha256,
                        NewValueSha256 = nestedResult.NewValueSha256,
                        SourceStartOffset = nestedResult.SourceStartOffset,
                        SourceEndOffset = nestedResult.SourceEndOffset,
                        OutputEndOffset = nestedResult.OutputEndOffset,
                    });
                }
            }
            var outputPackage = ReplaceParts(sourcePackage, replacements);
            var actualInnerChanges = ChangedParts(sourcePackage, outputPackage);
            var expectedInnerChanges = replacements.Keys.OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
            if (!actualInnerChanges.SequenceEqual(expectedInnerChanges, StringComparer.OrdinalIgnoreCase))
                throw new CodecException("presentation_edit_plan_scope_violation", $"Embedded workbook changed unexpected parts: {string.Join(", ", actualInnerChanges)}.", packagePath);
            if (patchedParts.ContainsKey(packagePath))
                throw new CodecException("presentation_edit_plan_scope_violation", "PPTX Edit Plan produced colliding outer part replacements.", packagePath);
            patchedParts.Add(packagePath, outputPackage);
        }
    }

    private static PptxXmlPatch[] CompileWorksheetValueXmlPatches(
        byte[] worksheetBytes,
        IReadOnlyList<PptxEditPlanProof> proofs)
    {
        var (xml, _) = DecodeXml(worksheetBytes);
        var prefixes = NamespacePrefixes(xml, SpreadsheetNamespace);
        if (prefixes.Count == 0)
            throw new CodecException("presentation_edit_target_missing", "Embedded worksheet does not declare the SpreadsheetML namespace.", proofs[0].MutationPartPath);
        var allRanges = NamespaceElementRanges(xml, null, prefixes);
        var cells = allRanges.Where(range => range.LocalName == "c").ToArray();
        var patches = new List<PptxXmlPatch>();
        foreach (var proof in proofs)
        {
            var operation = proof.Operation;
            var matchingCells = cells.Where(range => string.Equals(AttributeValue(xml, range, "r"), operation.EmbeddedCellReference, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matchingCells.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} embedded cell is missing or ambiguous.", operation.EmbeddedWorksheetPartPath);
            var values = allRanges.Where(range => range.LocalName == "v" && Contains(matchingCells[0], range)).ToArray();
            if (values.Length != 1)
                throw new CodecException("presentation_edit_target_mismatch", $"PPTX edit operation {operation.OperationId} embedded cell value is missing or ambiguous.", operation.EmbeddedWorksheetPartPath);
            var (start, end) = ElementTextSpan(xml, values[0]);
            if (xml[start..end] != EscapeText(operation.ExpectedValue))
                throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} embedded cell old value does not match.", operation.EmbeddedWorksheetPartPath);
            patches.Add(new PptxXmlPatch(operation, start, end, EscapeText(operation.Value), proof.SourceElementSha256, proof.MutationPartPath));
        }
        return OrderedNonOverlapping(patches, proofs[0].MutationPartPath);
    }

    private static PptxXmlPatch[] OrderedNonOverlapping(IEnumerable<PptxXmlPatch> patches, string partPath)
    {
        var ordered = patches.OrderBy(patch => patch.Start).ToArray();
        for (var index = 1; index < ordered.Length; index++)
            if (ordered[index - 1].End > ordered[index].Start)
                throw new CodecException("overlapping_presentation_edit_operations", "PPTX edit operations overlap in source XML.", partPath);
        return ordered;
    }

    private static HashSet<string> NamespacePrefixes(string xml, string namespaceUri)
    {
        var declarations = NamespacePattern().Matches(xml).Cast<Match>()
            .Select(match => (Prefix: match.Groups["prefix"].Value, Uri: match.Groups["uri"].Value))
            .ToList();
        declarations.AddRange(DefaultNamespacePattern().Matches(xml).Cast<Match>()
            .Select(match => (Prefix: string.Empty, Uri: match.Groups["uri"].Value)));
        return declarations.GroupBy(item => item.Prefix, StringComparer.Ordinal)
            .Where(group => group.All(item => item.Uri == namespaceUri))
            .Select(group => group.Key)
            .ToHashSet(StringComparer.Ordinal);
    }

    private static IReadOnlyList<XmlRange> NamespaceElementRanges(
        string xml,
        string? localName,
        IReadOnlySet<string> prefixes)
    {
        var ranges = new List<XmlRange>();
        var stack = new Stack<QualifiedOpenElement>();
        foreach (Match token in XmlTokenPattern().Matches(xml))
        {
            var value = token.Value;
            if (value.StartsWith("<!--", StringComparison.Ordinal) || value.StartsWith("<![CDATA[", StringComparison.Ordinal) || value.StartsWith("<?", StringComparison.Ordinal)) continue;
            var match = QualifiedTagPattern().Match(value);
            if (!match.Success) continue;
            var prefix = match.Groups["prefix"].Value;
            var name = match.Groups["name"].Value;
            var closing = value.StartsWith("</", StringComparison.Ordinal);
            var selfClosing = value.EndsWith("/>", StringComparison.Ordinal);
            if (!closing)
            {
                if (selfClosing)
                {
                    if ((localName is null || name == localName) && prefixes.Contains(prefix))
                        ranges.Add(new XmlRange(token.Index, token.Index + token.Length, name));
                }
                else stack.Push(new QualifiedOpenElement(prefix, name, token.Index));
                continue;
            }
            if (stack.Count == 0) throw new CodecException("invalid_presentation_xml", "PPTX Edit Plan encountered an unbalanced XML close tag.");
            var open = stack.Pop();
            if (open.Prefix != prefix || open.Name != name)
                throw new CodecException("invalid_presentation_xml", "PPTX Edit Plan encountered mismatched XML tags.");
            if ((localName is null || name == localName) && prefixes.Contains(prefix))
                ranges.Add(new XmlRange(open.Start, token.Index + token.Length, name));
        }
        if (stack.Count != 0) throw new CodecException("invalid_presentation_xml", "PPTX Edit Plan encountered unclosed XML tags.");
        return ranges.OrderBy(range => range.Start).ToArray();
    }

    private static bool Contains(XmlRange parent, XmlRange child) =>
        child.Start >= parent.Start && child.End <= parent.End && (child.Start != parent.Start || child.End != parent.End);

    private static string? AttributeValue(string xml, XmlRange range, string localName)
    {
        var tag = XmlTokenPattern().Matches(xml[range.Start..range.End]).Cast<Match>()
            .FirstOrDefault(match => !match.Value.StartsWith("</", StringComparison.Ordinal));
        if (tag is null) return null;
        var values = XmlAttributePattern().Matches(tag.Value).Cast<Match>()
            .Where(match => LocalAttributeName(match.Groups["name"].Value) == localName)
            .Select(match => match.Groups["value"].Value)
            .ToArray();
        return values.Length == 1 ? values[0] : null;
    }

    private static (int Start, int End) ElementTextSpan(string xml, XmlRange range)
    {
        var fragment = xml[range.Start..range.End];
        var openEnd = fragment.IndexOf('>');
        var closeStart = fragment.LastIndexOf("</", StringComparison.Ordinal);
        if (openEnd < 0 || closeStart <= openEnd || fragment[(openEnd + 1)..closeStart].Contains('<', StringComparison.Ordinal))
            throw new CodecException("presentation_edit_target_mismatch", "PPTX Edit Plan expected one plain XML text value.");
        return (range.Start + openEnd + 1, range.Start + closeStart);
    }

    private static string ElementText(string xml, XmlRange range)
    {
        var (start, end) = ElementTextSpan(xml, range);
        return xml[start..end];
    }

    private static bool HasEmbeddedWorkbookBinding(PresentationEditOperation operation) =>
        !string.IsNullOrEmpty(operation.EmbeddedPackagePartPath) ||
        !string.IsNullOrEmpty(operation.ExpectedEmbeddedPackageSha256) ||
        !string.IsNullOrEmpty(operation.EmbeddedPackageRelationshipId) ||
        !string.IsNullOrEmpty(operation.EmbeddedWorksheetPartPath) ||
        !string.IsNullOrEmpty(operation.ExpectedEmbeddedWorksheetSha256) ||
        !string.IsNullOrEmpty(operation.EmbeddedCellReference) ||
        !string.IsNullOrEmpty(operation.ChartFormula);

    private static bool HasDiagramBinding(PresentationEditOperation operation) =>
        !string.IsNullOrEmpty(operation.DiagramModelId) || operation.DiagramRunIndex != 0;

    private static void ValidateEmbeddedWorkbookBinding(PresentationEditOperation operation)
    {
        if (!EmbeddedPackagePartPathPattern().IsMatch(operation.EmbeddedPackagePartPath) || operation.EmbeddedPackagePartPath.Contains("..", StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid embedded package path.");
        if (!IsSha256(operation.ExpectedEmbeddedPackageSha256))
            throw new CodecException("invalid_presentation_edit_precondition", $"PPTX edit operation {operation.OperationId} has an invalid embedded package SHA-256.");
        if (string.IsNullOrWhiteSpace(operation.EmbeddedPackageRelationshipId) || operation.EmbeddedPackageRelationshipId.Length > 255)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid embedded package relationship ID.");
        if (!EmbeddedWorksheetPartPathPattern().IsMatch(operation.EmbeddedWorksheetPartPath) || operation.EmbeddedWorksheetPartPath.Contains("..", StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid embedded worksheet path.");
        if (!IsSha256(operation.ExpectedEmbeddedWorksheetSha256))
            throw new CodecException("invalid_presentation_edit_precondition", $"PPTX edit operation {operation.OperationId} has an invalid embedded worksheet SHA-256.");
        if (!CellReferencePattern().IsMatch(operation.EmbeddedCellReference))
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid embedded cell reference.");
        if (string.IsNullOrWhiteSpace(operation.ChartFormula) || operation.ChartFormula.Length > 1024)
            throw new CodecException("invalid_presentation_edit_target", $"PPTX edit operation {operation.OperationId} has an invalid chart formula binding.");
    }

    private static string LeafIndexKey(PresentationEditOperation operation) =>
        LeafKind(operation) switch
        {
            "chartDataValue" => $"{operation.ChartSeriesIndex}:{operation.ChartPointIndex}",
            "diagramText" => $"{operation.DiagramModelId}:{operation.DiagramRunIndex}",
            "paragraphAlignment" or "verticalAnchor" or "textBodyInsetLeftEmu" or "textBodyInsetTopEmu" or "textBodyInsetRightEmu" or "textBodyInsetBottomEmu" or "textBodyWrap" or "textBodyColumnCount" or "textBodyAutoFit" or "textBodyColumnDirection" or "fillRgb" or "fillScheme" or "lineRgb" or "lineScheme" or "lineWidthEmu" => $"native:{operation.NativeLeafIndex}",
            _ => operation.TextLeafIndex.ToString(System.Globalization.CultureInfo.InvariantCulture),
        };

    private static bool ValidFiniteNumber(string value) =>
        value.Length is > 0 and <= 128 && FiniteNumberPattern().IsMatch(value) &&
        double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number) &&
        double.IsFinite(number);
}
