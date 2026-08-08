using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeKit.Artifact.Wire.V1;
using SpreadsheetExtension = DocumentFormat.OpenXml.Spreadsheet.Extension;

namespace OfficeKit.Codec;

// Dynamic arrays use the same worksheet <f t="array"> shape as legacy array
// formulas, with distinguishing state in the workbook CellMetadataPart.
// OfficeKit owns one deliberately small source-free authoring profile: each
// requested anchor gets one XLDAPR metadata record and the exact range stays
// bound to the formula cell. Imported graphs remain byte-preserving and
// source-bound; this codec never edits or detaches their metadata.
internal sealed class XlsxDynamicArrayCodec
{
    private const string MetadataName = "XLDAPR";
    private const string ExtensionUri = "{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}";
    private const string DynamicArrayNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";

    private readonly WorkbookPart _workbookPart;
    private readonly HashSet<uint> _dynamicIndexes = [];
    private readonly List<Cell> _sourceFreeCells = [];

    internal XlsxDynamicArrayCodec(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart;
        ReadIndexes();
    }

    internal bool IsDynamic(Cell cell) =>
        cell.CellMetaIndex?.Value is { } index && _dynamicIndexes.Contains(index);

    internal bool HasOpaqueMetadata(Cell cell) =>
        cell.CellMetaIndex?.Value is { } index && !_dynamicIndexes.Contains(index);

    internal void ConfigureNewCell(Cell cell, CellArtifact source, bool sourceBound)
    {
        if (source.FormulaMetadata?.Kind != CellFormulaKind.DynamicArray) return;
        if (sourceBound)
            throw Unsupported(source, "cannot add a dynamic array to a source-bound workbook because imported dynamic arrays are read-only");
        ConfigureNewSourceFreeCell(cell, source);
    }

    internal void ConfigureNewSourceFreeCell(Cell cell, CellArtifact source)
    {
        if (source.FormulaMetadata?.Kind != CellFormulaKind.DynamicArray) return;
        if (cell.CellFormula?.FormulaType?.Value != CellFormulaValues.Array)
            throw Unsupported(source, "dynamic array requires an array worksheet formula");
        if (cell.CellMetaIndex is not null)
            throw Unsupported(source, "source-free dynamic array must not carry an existing cell metadata index");
        _sourceFreeCells.Add(cell);
    }

    internal void FinalizeSourceFree()
    {
        if (_sourceFreeCells.Count == 0) return;
        if (_workbookPart.CellMetadataPart is not null)
            throw new CodecException("invalid_cell_metadata", "Source-free dynamic-array authoring cannot replace an existing CellMetadataPart.");

        var metadataType = new MetadataType
        {
            Name = MetadataName,
            MinSupportedVersion = 120_000U,
            Copy = true,
            PasteAll = true,
            PasteValues = true,
            Merge = true,
            SplitFirst = true,
            RowColumnShift = true,
            ClearFormats = true,
            ClearComments = true,
            Assign = true,
            Coerce = true,
            CellMeta = true,
        };
        var extension = new SpreadsheetExtension(
            new DocumentFormat.OpenXml.Office2019.Excel.DynamicArray.DynamicArrayProperties
            {
                FDynamic = true,
                FCollapsed = false,
            })
        {
            Uri = ExtensionUri,
        };
        var future = new FutureMetadata(new FutureMetadataBlock(new ExtensionList(extension)))
        {
            Name = MetadataName,
            Count = 1U,
        };
        var cellMetadata = new CellMetadata { Count = checked((uint)_sourceFreeCells.Count) };
        for (var index = 0; index < _sourceFreeCells.Count; index++)
        {
            // CellMetaIndex is one-based while MetadataRecord.Val points to
            // the zero-based future-metadata block that describes the kind.
            _sourceFreeCells[index].CellMetaIndex = checked((uint)index + 1);
            cellMetadata.Append(new MetadataBlock(new MetadataRecord { TypeIndex = 1U, Val = 0U }));
        }

        var part = _workbookPart.AddNewPart<CellMetadataPart>();
        part.Metadata = new Metadata(
            new MetadataTypes(metadataType) { Count = 1U },
            future,
            cellMetadata);
        part.Metadata.Save();
    }

    internal void ApplyFormulaMetadata(Cell cell, CellArtifact desired, bool sourceBound)
    {
        var desiredDynamic = desired.FormulaMetadata?.Kind == CellFormulaKind.DynamicArray;
        var currentIndex = cell.CellMetaIndex?.Value;
        if (currentIndex is not null && !_dynamicIndexes.Contains(currentIndex.Value))
            throw Unsupported(desired, "formula is linked to opaque cell metadata and cannot be edited");

        var currentDynamic = currentIndex is not null;
        if (!currentDynamic)
        {
            if (desiredDynamic)
                throw Unsupported(
                    desired,
                    sourceBound
                        ? "cannot add a dynamic array because imported dynamic-array metadata is read-only"
                        : "cannot author a source-free dynamic array because the XLDAPR metadata graph is read-only");
            return;
        }

        if (!desiredDynamic)
            throw Unsupported(desired, "cannot detach an imported dynamic array because it is read-only");
        if (!FormulaMatches(cell.CellFormula, desired))
            throw Unsupported(desired, "cannot edit an imported dynamic-array formula because it is read-only");
    }

    private void ReadIndexes()
    {
        var metadata = _workbookPart.CellMetadataPart?.Metadata;
        if (metadata is null) return;
        var types = metadata.GetFirstChild<MetadataTypes>()?.Elements<MetadataType>().ToArray() ?? [];
        var dynamicTypes = types
            .Select((item, index) => (Item: item, Index: checked((uint)index + 1)))
            .Where(item => string.Equals(item.Item.Name?.Value, MetadataName, StringComparison.Ordinal) && item.Item.CellMeta?.Value == true)
            .ToArray();
        if (dynamicTypes.Length != 1) return;
        var dynamicTypeIndex = dynamicTypes[0].Index;

        var future = metadata.Elements<FutureMetadata>()
            .Where(item => string.Equals(item.Name?.Value, MetadataName, StringComparison.Ordinal))
            .ToArray();
        if (future.Length != 1) return;
        var futureBlocks = future[0].Elements<FutureMetadataBlock>().ToArray();
        if (future[0].Count?.Value != (uint)futureBlocks.Length) return;
        var dynamicValues = futureBlocks
            .Select((block, index) => (Block: block, Index: checked((uint)index)))
            .Where(item => IsDynamicBlock(item.Block))
            .Select(item => item.Index)
            .ToHashSet();
        if (dynamicValues.Count == 0) return;

        var cellMetadata = metadata.Elements<CellMetadata>().ToArray();
        if (cellMetadata.Length != 1) return;
        var blocks = cellMetadata[0].Elements<MetadataBlock>().ToArray();
        if (cellMetadata[0].Count?.Value != (uint)blocks.Length) return;
        for (var index = 0; index < blocks.Length; index++)
        {
            var records = blocks[index].Elements<MetadataRecord>().ToArray();
            var dynamicRecords = records.Where(record =>
                record.TypeIndex?.Value == dynamicTypeIndex &&
                record.Val?.Value is { } value && dynamicValues.Contains(value)).ToArray();
            if (dynamicRecords.Length == 0) continue;
            var cellIndex = checked((uint)index + 1);
            _dynamicIndexes.Add(cellIndex);
        }
    }

    private static bool FormulaMatches(CellFormula? current, CellArtifact desired)
    {
        if (current is null || current.FormulaType?.Value != CellFormulaValues.Array ||
            desired.FormulaMetadata is not { Kind: CellFormulaKind.DynamicArray } metadata)
            return false;
        var desiredBody = desired.Formula.Trim();
        if (desiredBody.StartsWith('=') ) desiredBody = desiredBody[1..];
        return string.Equals(current.Text ?? string.Empty, desiredBody, StringComparison.Ordinal) &&
               string.Equals(current.Reference?.Value ?? string.Empty, metadata.Reference, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDynamicBlock(FutureMetadataBlock block)
    {
        var matches = block.Descendants<SpreadsheetExtension>()
            .Where(extension => string.Equals(extension.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase))
            .SelectMany(extension => extension.Descendants())
            .Where(element =>
                element.LocalName == "dynamicArrayProperties" &&
                element.NamespaceUri == DynamicArrayNamespace &&
                Boolean(element.GetAttribute("fDynamic", string.Empty).Value) == true &&
                Boolean(element.GetAttribute("fCollapsed", string.Empty).Value) != true)
            .ToArray();
        return matches.Length == 1;
    }

    private static bool? Boolean(string? value) => value?.Trim().ToLowerInvariant() switch
    {
        "1" or "true" => true,
        "0" or "false" or "" => false,
        null => null,
        _ => null,
    };

    private static CodecException Unsupported(CellArtifact cell, string message) =>
        new(
            "unsupported_dynamic_array_edit",
            $"Cell {CellReference(cell.Row, cell.Column)} {message}.",
            CellReference(cell.Row, cell.Column));

    private static string CellReference(uint row, uint column)
    {
        var number = checked((int)column + 1);
        Span<char> buffer = stackalloc char[3];
        var position = buffer.Length;
        while (number > 0)
        {
            number--;
            buffer[--position] = (char)('A' + number % 26);
            number /= 26;
        }
        return $"{new string(buffer[position..])}{row + 1}";
    }
}
