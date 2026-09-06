using DocumentFormat.OpenXml;
using System.Globalization;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns a bounded, source-preserving DrawingML table projection. Table
// topology and merge ranges remain fixed after import; name, complete outer
// frame, non-visible title/description, recognized table-property flags,
// fixed-topology direct text runs, the bounded direct cell paint/border
// profile, and the bounded direct run-text-style profile are the source-bound
// edits. Recognized PowerPoint style/extension shells stay in the source graph
// instead of being rebuilt.
internal static class PptxTableCodec
{
    private const string TableGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/table";
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const int MaxColumns = 256;
    private const int MaxRows = 2_048;
    private const int MaxCellTextLength = 32_767;

    private readonly record struct MergeCellPlan(
        bool IsOrigin,
        int RowSpan,
        int ColumnSpan,
        bool HorizontalMerge,
        bool VerticalMerge);

    private readonly record struct NativeMergeCell(
        int RowSpan,
        int ColumnSpan,
        bool HorizontalMerge,
        bool VerticalMerge,
        bool HasRowSpan,
        bool HasColumnSpan,
        bool HasHorizontalMerge,
        bool HasVerticalMerge);

    internal static bool TryRead(P.GraphicFrame source, out PresentationTable table) =>
        TryRead(source, context: null, out table);

    internal static bool TryRead(P.GraphicFrame source, PptxPartContext? context, out PresentationTable table)
    {
        table = new PresentationTable();
        try
        {
            if (source.ChildElements.Count != 3 ||
                source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is not { Id.HasValue: true, Name: not null } ||
                source.Transform is not { } transform ||
                !TryReadFrame(transform, out var left, out var top, out var width, out var height, out var frameTransform) ||
                source.Graphic is not { ChildElements.Count: 1 } graphic ||
                graphic.GraphicData is not { ChildElements.Count: 1 } graphicData ||
                !string.Equals(graphicData.Uri?.Value, TableGraphicDataUri, StringComparison.Ordinal) ||
                graphicData.GetFirstChild<A.Table>() is not { } nativeTable ||
                nativeTable.ChildElements.Any(child => child is not A.TableProperties and not A.TableGrid and not A.TableRow))
                return false;

            var properties = nativeTable.Elements<A.TableProperties>().SingleOrDefault();
            var grid = nativeTable.Elements<A.TableGrid>().SingleOrDefault();
            var rows = nativeTable.Elements<A.TableRow>().ToArray();
            if (properties is null || grid is null ||
                nativeTable.Elements<A.TableProperties>().Count() != 1 ||
                nativeTable.Elements<A.TableGrid>().Count() != 1 ||
                nativeTable.ChildElements[0] is not A.TableProperties ||
                nativeTable.ChildElements[1] is not A.TableGrid ||
                !TablePropertiesSupported(properties) ||
                rows.Length is < 1 or > MaxRows)
                return false;

            var columns = grid.Elements<A.GridColumn>().ToArray();
            if (columns.Length is < 1 or > MaxColumns ||
                grid.ChildElements.Any(child => child is not A.GridColumn) ||
                columns.Any(column => column.Width?.Value is null or <= 0 || !GridColumnSupported(column)))
                return false;

            var result = new PresentationTable
            {
                LeftEmu = left,
                TopEmu = top,
                WidthEmu = width,
                HeightEmu = height,
                FrameTransform = frameTransform,
                Accessibility = PptxNonVisualAccessibilityCodec.Read(source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties),
            };
            result.ColumnWidthsEmu.Add(columns.Select(column => column.Width!.Value));
            if (properties.FirstRow is not null) result.FirstRow = properties.FirstRow.Value;
            if (properties.BandRow is not null) result.BandedRows = properties.BandRow.Value;
            if (properties.BandColumn is not null) result.BandedColumns = properties.BandColumn.Value;
            if (properties.FirstColumn is not null) result.FirstColumn = properties.FirstColumn.Value;
            if (properties.LastColumn is not null) result.LastColumn = properties.LastColumn.Value;
            if (properties.LastRow is not null) result.LastRow = properties.LastRow.Value;

            var nativeCells = new List<A.TableCell[]>(rows.Length);
            var cellStyleEditable = true;
            var cellTextStyleEditable = true;
            foreach (var nativeRow in rows)
            {
                if (nativeRow.Height?.Value is null or <= 0 || !TableRowSupported(nativeRow))
                    return false;
                var cells = nativeRow.Elements<A.TableCell>().ToArray();
                if (cells.Length != columns.Length)
                    return false;
                nativeCells.Add(cells);
                var row = new PresentationTableRow { HeightEmu = nativeRow.Height.Value };
                foreach (var cell in cells)
                {
                    if (!TryReadCell(cell, context, out var text, out var textBody, out var fill, out var borders, out var textStyle, out var styleEditable, out var textStyleEditable))
                        return false;
                    cellStyleEditable &= styleEditable;
                    cellTextStyleEditable &= textStyleEditable;
                    var modeledCell = new PresentationTableCell { Text = text };
                    if (textBody is not null) modeledCell.TextBody = textBody;
                    if (fill is not null) modeledCell.Fill = fill;
                    if (borders is not null) modeledCell.Borders = borders;
                    if (textStyle is not null) modeledCell.TextStyle = textStyle;
                    row.Cells.Add(modeledCell);
                }
                result.Rows.Add(row);
            }

            if (!TryReadMergeRanges(nativeCells, result)) return false;
            result.CellStyleEditable = cellStyleEditable;
            result.CellTextStyleEditable = cellTextStyleEditable;
            // PowerPoint commonly stores a table in its own coordinate space
            // and scales the graphic frame around it. Keep both dimensions
            // exactly as authored instead of rejecting the table or rewriting
            // its grid during a source-bound edit.
            if (!ScaledExtentSupported(width, result.ColumnWidthsEmu.Sum()) ||
                !ScaledExtentSupported(height, result.Rows.Sum(row => row.HeightEmu))) return false;
            table = result;
            return true;
        }
        catch (Exception error) when (error is InvalidOperationException or OverflowException)
        {
            table = new PresentationTable();
            return false;
        }
    }

    internal static P.GraphicFrame Build(PresentationElement element, uint nativeId, PptxPartContext slideContext)
    {
        var table = element.Table;
        Validate(table, element.Id, assets: slideContext.Assets);
        var mergePlan = CreateMergePlan(table, element.Id);
        var properties = new A.TableProperties();
        if (table.HasFirstRow) properties.FirstRow = table.FirstRow;
        if (table.HasBandedRows) properties.BandRow = table.BandedRows;
        if (table.HasBandedColumns) properties.BandColumn = table.BandedColumns;
        if (table.HasFirstColumn) properties.FirstColumn = table.FirstColumn;
        if (table.HasLastColumn) properties.LastColumn = table.LastColumn;
        if (table.HasLastRow) properties.LastRow = table.LastRow;
        var grid = new A.TableGrid();
        foreach (var width in table.ColumnWidthsEmu) grid.Append(new A.GridColumn { Width = width });
        var nativeTable = new A.Table(properties, grid);
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var sourceRow = table.Rows[rowIndex];
            var row = new A.TableRow { Height = sourceRow.HeightEmu };
            for (var columnIndex = 0; columnIndex < sourceRow.Cells.Count; columnIndex++)
            {
                var sourceCell = sourceRow.Cells[columnIndex];
                var cell = BuildCell(sourceCell, table.HasFirstRow && table.FirstRow && rowIndex == 0, table, slideContext);
                if (mergePlan.TryGetValue((rowIndex, columnIndex), out var merge))
                {
                    if (merge.IsOrigin)
                    {
                        if (merge.RowSpan > 1) cell.RowSpan = merge.RowSpan;
                        if (merge.ColumnSpan > 1) cell.GridSpan = merge.ColumnSpan;
                    }
                    else
                    {
                        if (merge.HorizontalMerge) cell.HorizontalMerge = true;
                        if (merge.VerticalMerge) cell.VerticalMerge = true;
                    }
                }
                row.Append(cell);
            }
            nativeTable.Append(row);
        }
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = element.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, table.Accessibility);
        var transform = new P.Transform(
            new A.Offset { X = table.LeftEmu, Y = table.TopEmu },
            new A.Extents { Cx = table.WidthEmu, Cy = table.HeightEmu });
        PptxFrameTransformCodec.Apply(transform, table.FrameTransform);
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                nonVisual,
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            transform,
            new A.Graphic(new A.GraphicData(nativeTable) { Uri = TableGraphicDataUri }));
    }

    internal static void Apply(P.GraphicFrame source, PresentationElement requested, PptxPartContext? slideContext = null)
    {
        if (!TryRead(source, slideContext, out var original))
            throw new CodecException("unsupported_presentation_edit", $"Presentation table {requested.Id} no longer matches the editable table profile.");
        ValidateRequest(original, requested);
        var table = requested.Table;
        PptxNonVisualAccessibilityCodec.ApplyBound(
            source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
            table.Accessibility,
            "table");
        source.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Name = requested.Name;
        SetFrame(source.Transform!, table);
        var nativeTable = source.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        var properties = nativeTable.GetFirstChild<A.TableProperties>()!;
        if (table.HasFirstRow) properties.FirstRow = table.FirstRow;
        else properties.FirstRow = null;
        if (table.HasBandedRows) properties.BandRow = table.BandedRows;
        else properties.BandRow = null;
        if (table.HasBandedColumns) properties.BandColumn = table.BandedColumns;
        else properties.BandColumn = null;
        if (table.HasFirstColumn) properties.FirstColumn = table.FirstColumn;
        else properties.FirstColumn = null;
        if (table.HasLastColumn) properties.LastColumn = table.LastColumn;
        else properties.LastColumn = null;
        if (table.HasLastRow) properties.LastRow = table.LastRow;
        else properties.LastRow = null;
        var columns = nativeTable.GetFirstChild<A.TableGrid>()!.Elements<A.GridColumn>().ToArray();
        for (var index = 0; index < columns.Length; index++) columns[index].Width = table.ColumnWidthsEmu[index];
        var rows = nativeTable.Elements<A.TableRow>().ToArray();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            rows[rowIndex].Height = table.Rows[rowIndex].HeightEmu;
            var cells = rows[rowIndex].Elements<A.TableCell>().ToArray();
            for (var columnIndex = 0; columnIndex < cells.Length; columnIndex++)
            {
                var requestedText = table.Rows[rowIndex].Cells[columnIndex].Text;
                var sourceText = CellText(cells[columnIndex]);
                PatchCellProperties(cells[columnIndex], table.Rows[rowIndex].Cells[columnIndex], slideContext);
                PatchCellTextBody(cells[columnIndex], table.Rows[rowIndex].Cells[columnIndex], slideContext);
                PatchCellTextStyle(cells[columnIndex], table.Rows[rowIndex].Cells[columnIndex], slideContext);
                // A structured body patch already updates each fixed inline,
                // including cached field text. Do not run the legacy plain
                // text distributor afterwards: it only owns direct a:r/a:t
                // leaves and must never flatten a field into ordinary text.
                if (table.Rows[rowIndex].Cells[columnIndex].TextBody is not null)
                    continue;
                // Text replacement keeps the source paragraph and run
                // topology. A direct plain-text paragraph may contain one or
                // more runs; new text is distributed deterministically across
                // those existing leaves without rebuilding the cell body.
                if (string.Equals(requestedText, sourceText, StringComparison.Ordinal))
                    continue;
                var textRunLeaves = TextRunLeaves(cells[columnIndex]);
                if (textRunLeaves is null)
                {
                    if (requestedText.Length != 0)
                        throw new CodecException("unsupported_presentation_edit", $"Presentation table {requested.Id} cannot add text to an empty or covered cell without a source text leaf.");
                    continue;
                }
                var lines = requestedText.Split('\n');
                if (lines.Length != textRunLeaves.Count)
                    throw new CodecException("unsupported_presentation_edit", $"Presentation table {requested.Id} must preserve the source paragraph count when editing a multi-paragraph cell.");
                for (var lineIndex = 0; lineIndex < textRunLeaves.Count; lineIndex++)
                    SetParagraphText(textRunLeaves[lineIndex], lines[lineIndex]);
            }
        }
    }

    internal static void Validate(
        PresentationTable? table,
        string elementId,
        bool allowScaledFrame = false,
        PptxAssetCatalog? assets = null)
    {
        if (table is null) throw Invalid(elementId, "payload is missing");
        if (table.LeftEmu < 0 || table.TopEmu < 0 || table.WidthEmu <= 0 || table.HeightEmu <= 0)
            throw Invalid(elementId, "frame must have non-negative coordinates and positive dimensions");
        PptxFrameTransformCodec.Validate(table.FrameTransform, elementId, "table");
        if (table.ColumnWidthsEmu.Count is < 1 or > MaxColumns || table.Rows.Count is < 1 or > MaxRows)
            throw Invalid(elementId, $"grid must contain 1-{MaxColumns} columns and 1-{MaxRows} rows");
        if (table.ColumnWidthsEmu.Any(width => width <= 0) ||
            (!allowScaledFrame && Sum(table.ColumnWidthsEmu, elementId) != table.WidthEmu) ||
            (allowScaledFrame && !ScaledExtentSupported(table.WidthEmu, Sum(table.ColumnWidthsEmu, elementId))))
            throw Invalid(elementId, "positive column widths must fit the outer frame width");
        if (table.Rows.Any(row => row.HeightEmu <= 0 || row.Cells.Count != table.ColumnWidthsEmu.Count) ||
            (!allowScaledFrame && Sum(table.Rows.Select(row => row.HeightEmu), elementId) != table.HeightEmu) ||
            (allowScaledFrame && !ScaledExtentSupported(table.HeightEmu, Sum(table.Rows.Select(row => row.HeightEmu), elementId))))
            throw Invalid(elementId, "positive row heights must fit the outer frame height and every row must match the grid width");
        foreach (var cell in table.Rows.SelectMany(row => row.Cells))
        {
            if (cell.Text.Length > MaxCellTextLength || cell.Text.Any(character => char.IsControl(character) && character is not '\t' and not '\n' and not '\r'))
                throw Invalid(elementId, $"cell text must contain at most {MaxCellTextLength} characters and no unsupported controls");
            if (cell.TextBody is not null)
            {
                if (!cell.Text.Equals(PptxTextCodec.Flatten(cell.TextBody), StringComparison.Ordinal))
                    throw Invalid(elementId, "cell text must equal its structured text_body content");
                var validationShape = new PresentationShape { Text = cell.Text, TextBody = cell.TextBody.Clone() };
                PptxTextCodec.Validate(validationShape);
            }
            ValidateCellFill(cell.Fill, elementId, assets);
            ValidateCellBorders(cell.Borders, elementId);
            ValidateCellTextStyle(cell.TextStyle, elementId);
        }
        PptxNonVisualAccessibilityCodec.Validate(table.Accessibility, elementId, "table");
        if (table.DefaultCellFillCase == PresentationTable.DefaultCellFillOneofCase.DefaultCellFillRgb)
            PptxColor.Normalize(table.DefaultCellFillRgb);
        if (table.DefaultCellFillCase == PresentationTable.DefaultCellFillOneofCase.NoDefaultCellFill && !table.NoDefaultCellFill)
            throw Invalid(elementId, "no_default_cell_fill must be true when selected");
        if (table.DefaultTextStyle is not null)
        {
            if (table.DefaultTextStyle.HasFontSizePoints && table.DefaultTextStyle.FontSizePoints is <= 0 or > 1000)
                throw Invalid(elementId, "default text font size must be from 0 through 1000 points");
            if (table.DefaultTextStyle.HasColorRgb) PptxColor.Normalize(table.DefaultTextStyle.ColorRgb);
            if (table.DefaultTextStyle.HasColorScheme) PptxColor.NormalizeScheme(table.DefaultTextStyle.ColorScheme);
            if (table.DefaultTextStyle.HasColorOpacityThousandthPercent &&
                table.DefaultTextStyle.ColorCase == PresentationTextStyle.ColorOneofCase.None)
                throw Invalid(elementId, "default text color opacity requires a modeled color");
            if (table.DefaultTextStyle.HasColorOpacityThousandthPercent && table.DefaultTextStyle.ColorOpacityThousandthPercent > 100_000)
                throw Invalid(elementId, "default text color opacity must be at most 100000 thousandths of a percent");
        }
        _ = CreateMergePlan(table, elementId);
    }

    internal static void ScrubModeledContent(P.GraphicFrame source, PptxPartContext? context = null)
    {
        // Direct cell paint and borders are part of the bounded source-bound
        // style profile only when every physical cell is representable. Keep
        // unsupported relationship/effect children in the residual hash.
        var parsed = TryRead(source, context, out var parsedTable);
        var scrubCellStyles = parsed && parsedTable.CellStyleEditable;
        var scrubCellTextStyles = parsed;
        if (source.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            PptxNonVisualAccessibilityCodec.ScrubModeledContent(nonVisual);
            nonVisual.Name = string.Empty;
        }
        if (source.Transform is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
            PptxFrameTransformCodec.Scrub(transform);
        }
        var table = source.Graphic?.GraphicData?.GetFirstChild<A.Table>();
        if (table is null) return;
        if (table.GetFirstChild<A.TableProperties>() is { } properties)
        {
            // These six flags are the only table-style leaves issued by the
            // source-bound PPJ profile. Keep table-style IDs, no-fill shells,
            // and extension children in the residual hash.
            properties.FirstRow = null;
            properties.BandRow = null;
            properties.BandColumn = null;
            properties.FirstColumn = null;
            properties.LastColumn = null;
            properties.LastRow = null;
        }
        foreach (var column in table.GetFirstChild<A.TableGrid>()?.Elements<A.GridColumn>() ?? []) column.Width = 1L;
        foreach (var row in table.Elements<A.TableRow>())
        {
            row.Height = 1L;
            foreach (var cell in row.Elements<A.TableCell>())
            {
                foreach (var text in (TextRunLeaves(cell)?.SelectMany(leaves => leaves) ?? TextLeaves(cell) ?? [])) text.Text = string.Empty;
                if (scrubCellStyles && cell.GetFirstChild<A.TableCellProperties>() is { } cellProperties)
                    foreach (var child in cellProperties.ChildElements.Where(child => IsCellFill(child) || IsCellBorder(child)).ToArray())
                        child.Remove();
                // Text-body editability is a table-wide capability in the
                // public model, but residual hashing must scrub each cell only
                // when that cell independently satisfies the bounded profile.
                // This keeps a modeled picture marker's relationship leaf
                // editable without masking an unrelated opaque cell.
                var scrubThisCellText = scrubCellTextStyles && IsCellTextBodyEditable(cell, context);
                if (scrubThisCellText && cell.GetFirstChild<A.TextBody>() is { } textBody)
                {
                    var temporary = new P.Shape { TextBody = new P.TextBody { InnerXml = textBody.InnerXml } };
                    PptxTextCodec.ScrubModeledContent(temporary.TextBody, context);
                    textBody.InnerXml = temporary.TextBody!.InnerXml;
                }
            }
        }
    }

    private static bool IsCellTextBodyEditable(A.TableCell cell, PptxPartContext? context)
    {
        return TryReadCell(cell, context, out _, out _, out _, out _, out _, out _, out var editable) && editable;
    }

    private static void ValidateRequest(PresentationTable original, PresentationElement requested)
    {
        var scaledFrame = original.ColumnWidthsEmu.Sum() != original.WidthEmu ||
            original.Rows.Sum(row => row.HeightEmu) != original.HeightEmu;
        Validate(requested.Table, requested.Id, allowScaledFrame: scaledFrame);
        if (requested.Name.Length > 1_024) throw Invalid(requested.Id, "name exceeds 1024 characters");
        var allowed = original.Clone();
        allowed.LeftEmu = requested.Table.LeftEmu;
        allowed.TopEmu = requested.Table.TopEmu;
        allowed.WidthEmu = requested.Table.WidthEmu;
        allowed.HeightEmu = requested.Table.HeightEmu;
        allowed.FrameTransform = requested.Table.FrameTransform?.Clone();
        allowed.ColumnWidthsEmu.Clear();
        allowed.ColumnWidthsEmu.Add(requested.Table.ColumnWidthsEmu);
        allowed.Accessibility = requested.Table.Accessibility?.Clone();
        if (requested.Table.HasFirstRow) allowed.FirstRow = requested.Table.FirstRow;
        else allowed.ClearFirstRow();
        if (requested.Table.HasBandedRows) allowed.BandedRows = requested.Table.BandedRows;
        else allowed.ClearBandedRows();
        if (requested.Table.HasBandedColumns) allowed.BandedColumns = requested.Table.BandedColumns;
        else allowed.ClearBandedColumns();
        if (requested.Table.HasFirstColumn) allowed.FirstColumn = requested.Table.FirstColumn;
        else allowed.ClearFirstColumn();
        if (requested.Table.HasLastColumn) allowed.LastColumn = requested.Table.LastColumn;
        else allowed.ClearLastColumn();
        if (requested.Table.HasLastRow) allowed.LastRow = requested.Table.LastRow;
        else allowed.ClearLastRow();
        for (var rowIndex = 0; rowIndex < allowed.Rows.Count; rowIndex++)
        {
            allowed.Rows[rowIndex].HeightEmu = requested.Table.Rows[rowIndex].HeightEmu;
            for (var columnIndex = 0; columnIndex < allowed.Rows[rowIndex].Cells.Count; columnIndex++)
            {
                var allowedCell = allowed.Rows[rowIndex].Cells[columnIndex];
                var requestedCell = requested.Table.Rows[rowIndex].Cells[columnIndex];
                var originalCell = original.Rows[rowIndex].Cells[columnIndex];
                allowedCell.Text = requestedCell.Text;
                if (original.CellStyleEditable)
                {
                    allowedCell.Fill = requestedCell.Fill?.Clone();
                    allowedCell.Borders = requestedCell.Borders?.Clone();
                }
                if (original.CellTextStyleEditable)
                {
                    allowedCell.TextStyle = requestedCell.TextStyle?.Clone();
                    if (originalCell.TextBody is not null)
                    {
                        if (requestedCell.TextBody is null)
                            throw new CodecException("unsupported_presentation_edit", "Source-preserving mixed-run table-cell text must retain its structured text body.");
                        if (requestedCell.TextStyle is not null)
                            throw new CodecException("unsupported_presentation_edit", "Source-preserving mixed-run table-cell text must use per-run styles inside text_body, not a uniform textStyle.");
                        allowedCell.TextBody = requestedCell.TextBody.Clone();
                    }
                }
            }
        }
        if (!allowed.Equals(requested.Table))
            throw new CodecException("unsupported_presentation_edit", $"Presentation table {requested.Id} may edit only its name, complete frame transform, alternative text, recognized table-property flags, bounded direct cell paint/borders, fixed-topology plain cell text, and bounded mixed-run text bodies.");
    }

    private static bool TryReadFrame(
        P.Transform transform,
        out long left,
        out long top,
        out long width,
        out long height,
        out PresentationFrameTransform? frameTransform)
    {
        left = top = width = height = 0;
        frameTransform = null;
        if (transform.ChildElements.Count != 2 || transform.Offset is not { } offset || transform.Extents is not { } extents ||
            !HasOnlyAttributes(offset, "x", "y") || !HasOnlyAttributes(extents, "cx", "cy") ||
            offset.X?.Value is null || offset.Y?.Value is null || extents.Cx?.Value is null or <= 0 || extents.Cy?.Value is null or <= 0 ||
            offset.X.Value < 0 || offset.Y.Value < 0 ||
            !PptxFrameTransformCodec.TryRead(transform, out frameTransform))
            return false;
        left = offset.X.Value;
        top = offset.Y.Value;
        width = extents.Cx.Value;
        height = extents.Cy.Value;
        return true;
    }

    private static bool TryReadCell(
        A.TableCell cell,
        PptxPartContext? context,
        out string text,
        out PresentationTextBody? textBody,
        out PresentationTableCellFill? fill,
        out PresentationTableCellBorders? borders,
        out PresentationTextStyle? textStyle,
        out bool styleEditable,
        out bool textStyleEditable)
    {
        text = string.Empty;
        textBody = null;
        fill = null;
        borders = null;
        textStyle = null;
        styleEditable = true;
        textStyleEditable = true;
        if (!HasOnlyAttributes(cell, "rowSpan", "gridSpan", "hMerge", "vMerge")) return false;
        // Covered merge cells are legal and intentionally have no text body.
        // They remain fixed-topology cells and cannot receive new text during
        // a source-bound edit.
        if (cell.ChildElements.Count == 0) return true;
        if (cell.Elements<A.TextBody>().Count() != 1 || cell.Elements<A.TableCellProperties>().Count() != 1) return false;
        if (!TryReadCellProperties(cell.GetFirstChild<A.TableCellProperties>(), context, out fill, out borders, out styleEditable))
            return false;
        var body = cell.GetFirstChild<A.TextBody>()!;
        if (body.ChildElements.Count < 3 || body.ChildElements[0] is not A.BodyProperties ||
            body.ChildElements[1] is not A.ListStyle || body.ChildElements.Skip(2).Any(child => child is not A.Paragraph))
            return false;

        // Keep the table model plain, but accept the common source-bound form
        // where one cell contains several paragraphs and each paragraph has
        // one or more direct text/field/line-break inlines. Joining those
        // paragraphs gives the Agent a single editable cell value without
        // flattening any run properties; Apply() splices the same
        // paragraph/inline leaves in place.
        var paragraphs = body.Elements<A.Paragraph>().ToArray();
        var lines = new List<string>(paragraphs.Length);
        var hasAnyInline = paragraphs.Any(paragraph => paragraph.ChildElements.Any(child => child is A.Run or A.Field or A.Break));
        if (!hasAnyInline)
        {
            // A normal empty cell has a paragraph shell but no source inline.
            // Keep the cell readable; adding text remains fail-closed.
            text = string.Empty;
            return true;
        }
        foreach (var paragraph in paragraphs)
        {
            var inlines = paragraph.ChildElements
                .Where(child => child is A.Run or A.Field or A.Break)
                .ToArray();
            if (inlines.Length < 1)
                return false;
            var line = new System.Text.StringBuilder();
            foreach (var inline in inlines)
            {
                if (inline is A.Break)
                {
                    line.Append('\n');
                    continue;
                }
                if (inline is not (A.Run or A.Field) || inline.Descendants<A.Text>().Count() != 1)
                    return false;
                var value = inline.GetFirstChild<A.Text>();
                if (value is null || value.Text.Contains('\n')) return false;
                line.Append(value.Text);
            }
            lines.Add(line.ToString());
        }
        text = string.Join("\n", lines);
        if (PptxTextCodec.SupportsEditing(body))
        {
            try
            {
                var semantic = PptxTextCodec.ReadDrawingTextBody(body, context);
                var runs = semantic.Paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
                var styles = runs.Select(TextStyle).ToArray();
                var hasStructuredInline = runs.Any(run => run.ContentCase is
                    PresentationTextRun.ContentOneofCase.Field or PresentationTextRun.ContentOneofCase.LineBreak);
                if (runs.Length > 0 && !hasStructuredInline && runs.All(run =>
                        run.ContentCase == PresentationTextRun.ContentOneofCase.Text &&
                        CellTextStyleSupported(run)) &&
                    styles.Skip(1).All(style => Equals(style, styles[0])) &&
                    semantic.Paragraphs.All(paragraph => !PptxParagraphPropertiesCodec.HasModeledProperties(paragraph)) &&
                    !PptxBodyPropertiesCodec.HasModeledProperties(semantic.BodyProperties))
                {
                    textStyle = styles[0];
                }
                else if (runs.Length > 0 && IsBoundedMixedRunTextBody(semantic))
                {
                    // Preserve a fixed-topology, direct-run body when the
                    // cell intentionally contains heterogeneous run styles.
                    // The projector can then expose per-run style leaves and
                    // the compiler can apply them without flattening to one
                    // uniform cell style.
                    textBody = semantic;
                }
                else
                {
                    textStyleEditable = false;
                }
            }
            catch (CodecException)
            {
                textStyleEditable = false;
            }
        }
        else if (paragraphs.Length > 0)
        {
            textStyleEditable = false;
        }
        return text.Length <= MaxCellTextLength;
    }

    internal static bool IsBoundedMixedRunTextBody(PresentationTextBody body) =>
        BoundedTableBodyProperties(body.BodyProperties) &&
        body.ListStyles.Count == 0 &&
        body.Paragraphs.All(paragraph =>
            BoundedTableParagraphProperties(paragraph) &&
            paragraph.Runs.All(run =>
                (run.ContentCase is PresentationTextRun.ContentOneofCase.Text or PresentationTextRun.ContentOneofCase.Field or PresentationTextRun.ContentOneofCase.LineBreak) &&
                CellTextStyleSupported(run)));

    // A structured table-cell text body may expose the direct a:bodyPr leaves
    // that have a stable PPJ textBoxStyle spelling. Unsupported bodyPr
    // choices (rotation, overflow, normAutofit percentages and explicit
    // delete markers) remain source-owned so a rich-text edit cannot silently
    // drop them. The unknown attributes/children themselves remain on the
    // native body and are preserved by PptxBodyPropertiesCodec.Apply.
    private static bool BoundedTableBodyProperties(PresentationTextBodyProperties? properties) =>
        PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(properties);

    // A source-bound table cell may carry direct paragraph alignment and
    // concrete spacing without requiring the broader list/layout/effects
    // profile used by ordinary text boxes. Keep this deliberately narrow:
    // these are stable DrawingML leaves, while complex bullet/effect graphs
    // and inherited defaults still need their own topology and cascade
    // contracts. A bounded direct a:bodyPr subset is accepted separately so
    // cell text can expose vertical anchoring, wrapping and insets without
    // flattening the text body.
    private static bool BoundedTableParagraphProperties(PresentationTextParagraph paragraph) =>
        (!paragraph.HasAlignment || paragraph.Alignment is "left" or "center" or "right" or "justify" or "distributed") &&
        paragraph.LeftMarginCase is PresentationTextParagraph.LeftMarginOneofCase.None or PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu &&
        paragraph.IndentationCase is PresentationTextParagraph.IndentationOneofCase.None or PresentationTextParagraph.IndentationOneofCase.IndentEmu &&
        BoundedTableSpacing(paragraph.LineSpacingCase, PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints, PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier) &&
        BoundedTableSpacing(paragraph.SpaceBeforeCase, PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints, PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier) &&
        BoundedTableSpacing(paragraph.SpaceAfterCase, PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints, PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier) &&
        (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.None ||
         paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.NoBullet ||
         paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.BulletCharacter ||
         paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.AutoNumber ||
         paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
         paragraph.PictureBullet.SourceCase is PresentationPictureBullet.SourceOneofCase.AssetId or
             PresentationPictureBullet.SourceOneofCase.Uri) &&
        BoundedTableBulletStyle(paragraph) &&
        paragraph.TabStops.Count <= 256 &&
        (!paragraph.HasNoTabStops || paragraph.NoTabStops) &&
        BoundedTableDefaultRunStyle(paragraph);

    private static bool BoundedTableSpacing<T>(T actual, T points, T multiplier)
        where T : Enum =>
        EqualityComparer<T>.Default.Equals(actual, default) ||
        EqualityComparer<T>.Default.Equals(actual, points) ||
        EqualityComparer<T>.Default.Equals(actual, multiplier);

    private static bool BoundedTableBulletStyle(PresentationTextParagraph paragraph)
    {
        var hasStyle = paragraph.BulletFontCase != PresentationTextParagraph.BulletFontOneofCase.None ||
            paragraph.BulletColorCase != PresentationTextParagraph.BulletColorOneofCase.None ||
            paragraph.BulletSizeCase != PresentationTextParagraph.BulletSizeOneofCase.None;
        if (!hasStyle) return true;
        if (paragraph.BulletCase is not (PresentationTextParagraph.BulletOneofCase.BulletCharacter or
            PresentationTextParagraph.BulletOneofCase.AutoNumber or
            PresentationTextParagraph.BulletOneofCase.PictureBullet)) return false;
        if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
            paragraph.PictureBullet.SourceCase is not (PresentationPictureBullet.SourceOneofCase.AssetId or
                PresentationPictureBullet.SourceOneofCase.Uri))
            return false;
        var font = paragraph.BulletFontCase is PresentationTextParagraph.BulletFontOneofCase.None or
            PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily or
            PresentationTextParagraph.BulletFontOneofCase.BulletFontFollowText;
        var color = paragraph.BulletColorCase is PresentationTextParagraph.BulletColorOneofCase.None or
            PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb or
            PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme or
            PresentationTextParagraph.BulletColorOneofCase.BulletColorFollowText;
        var size = paragraph.BulletSizeCase is PresentationTextParagraph.BulletSizeOneofCase.None or
            PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints or
            PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent or
            PresentationTextParagraph.BulletSizeOneofCase.BulletSizeFollowText;
        return font && color && size;
    }

    private static bool BoundedTableDefaultRunStyle(PresentationTextParagraph paragraph) =>
        paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.None ||
        paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
        paragraph.DefaultRunProperties is { } style &&
        !style.HasFontKerningPoints && !style.HasFontBaselinePercent && !style.HasFontSpacingPoints &&
        !style.HasFontCaps && style.HighlightCase == PresentationTextStyle.HighlightOneofCase.None &&
        style.GradientFill is null && style.Shadow is null &&
        style.ColorCase != PresentationTextStyle.ColorOneofCase.ColorScheme &&
        (!style.HasLanguage || style.Language.Equals("en-US", StringComparison.OrdinalIgnoreCase));

    private static PresentationTextStyle? TextStyle(PresentationTextRun run)
    {
        var style = new PresentationTextStyle();
        var hasStyle = false;
        if (run.HasBold) { style.Bold = run.Bold; hasStyle = true; }
        if (run.HasItalic) { style.Italic = run.Italic; hasStyle = true; }
        if (run.HasFontSizePoints) { style.FontSizePoints = run.FontSizePoints; hasStyle = true; }
        if (run.HasFontFamily) { style.FontFamily = run.FontFamily; hasStyle = true; }
        if (run.HasFontFamilyEastAsia) { style.FontFamilyEastAsia = run.FontFamilyEastAsia; hasStyle = true; }
        if (run.HasFontFamilyComplexScript) { style.FontFamilyComplexScript = run.FontFamilyComplexScript; hasStyle = true; }
        if (run.HasColorRgb) { style.ColorRgb = run.ColorRgb; hasStyle = true; }
        if (run.HasColorOpacityThousandthPercent) { style.ColorOpacityThousandthPercent = run.ColorOpacityThousandthPercent; hasStyle = true; }
        if (run.HasUnderline) { style.Underline = run.Underline; hasStyle = true; }
        if (run.HasStrike) { style.Strike = run.Strike; hasStyle = true; }
        return hasStyle ? style : null;
    }

    private static bool CellTextStyleSupported(PresentationTextRun run) =>
        !run.HasFontKerningPoints && !run.HasFontBaselinePercent && !run.HasFontSpacingPoints &&
        !run.HasFontCaps && run.HighlightCase == PresentationTextRun.HighlightOneofCase.None &&
        run.GradientFill is null && run.Shadow is null &&
        !run.HasColorScheme &&
        (!run.HasLanguage || run.Language.Equals("en-US", StringComparison.OrdinalIgnoreCase)) &&
        run.HyperlinkCase == PresentationTextRun.HyperlinkOneofCase.None;

    private static void ValidateCellTextStyle(PresentationTextStyle? style, string elementId)
    {
        if (style is null) return;
        if (style.HasFontSizePoints && (style.FontSizePoints <= 0 || style.FontSizePoints > 1000))
            throw Invalid(elementId, "cell text font size must be from 0 through 1000 points");
        if (style.HasFontFamily && string.IsNullOrWhiteSpace(style.FontFamily))
            throw Invalid(elementId, "cell text font family must not be empty");
        if (style.HasFontFamilyEastAsia && string.IsNullOrWhiteSpace(style.FontFamilyEastAsia))
            throw Invalid(elementId, "cell text East Asian font family must not be empty");
        if (style.HasFontFamilyComplexScript && string.IsNullOrWhiteSpace(style.FontFamilyComplexScript))
            throw Invalid(elementId, "cell text complex-script font family must not be empty");
        if (style.HasColorRgb) PptxColor.Normalize(style.ColorRgb);
        if (style.HasColorScheme) PptxColor.NormalizeScheme(style.ColorScheme);
        if (style.HasColorOpacityThousandthPercent &&
            style.ColorCase == PresentationTextStyle.ColorOneofCase.None)
            throw Invalid(elementId, "cell text color opacity requires a modeled color");
        if (style.HasColorOpacityThousandthPercent && style.ColorOpacityThousandthPercent > 100_000)
            throw Invalid(elementId, "cell text color opacity must be at most 100000 thousandths of a percent");
        if (style.HasUnderline) _ = PptxTextDecoration.NormalizeUnderline(style.Underline);
        if (style.HasStrike) _ = PptxTextDecoration.NormalizeStrike(style.Strike);
    }

    private static bool TryReadCellProperties(
        A.TableCellProperties? properties,
        PptxPartContext? context,
        out PresentationTableCellFill? fill,
        out PresentationTableCellBorders? borders,
        out bool styleEditable)
    {
        fill = null;
        borders = null;
        styleEditable = true;
        if (properties is null) return true;

        var fills = properties.ChildElements.Where(IsCellFill).ToArray();
        if (fills.Length > 1) return false;
        if (fills.SingleOrDefault() is { } nativeFill &&
            !TryReadCellFill(nativeFill, context, out fill))
            styleEditable = false;

        var parsedBorders = new PresentationTableCellBorders();
        var hasBorder = false;
        foreach (var (name, assign) in new (string Name, Action<SpreadsheetChartLineStyleArtifact?> Set)[]
        {
            ("lnL", value => parsedBorders.Left = value),
            ("lnT", value => parsedBorders.Top = value),
            ("lnR", value => parsedBorders.Right = value),
            ("lnB", value => parsedBorders.Bottom = value),
        })
        {
            var nativeBorder = properties.ChildElements.SingleOrDefault(child => child.LocalName == name && child.NamespaceUri == DrawingNamespace);
            if (properties.ChildElements.Count(child => child.LocalName == name && child.NamespaceUri == DrawingNamespace) > 1)
                return false;
            if (nativeBorder is null) continue;
            if (!TryReadCellBorder(nativeBorder, out var border))
            {
                styleEditable = false;
                continue;
            }
            if (border is not null) hasBorder = true;
            assign(border);
        }
        if (hasBorder) borders = parsedBorders;
        return true;
    }

    private static bool IsCellFill(OpenXmlElement child) =>
        child.NamespaceUri == DrawingNamespace &&
        child.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill";

    private static bool IsCellBorder(OpenXmlElement child) =>
        child.NamespaceUri == DrawingNamespace &&
        child.LocalName is "lnL" or "lnT" or "lnR" or "lnB";

    private static bool TryReadCellFill(
        OpenXmlElement source,
        PptxPartContext? context,
        out PresentationTableCellFill? fill)
    {
        fill = null;
        if (source is A.NoFill noFill)
        {
            if (noFill.GetAttributes().Count != 0 || noFill.ChildElements.Count != 0) return false;
            fill = new PresentationTableCellFill { NoFill = true };
            return true;
        }
        if (source is A.SolidFill solid &&
            PptxColor.TryDirectSolidRgbWithOpacity(solid, out var rgb, out var opacity))
        {
            fill = new PresentationTableCellFill { SolidRgb = rgb };
            if (opacity is { } alpha) fill.OpacityThousandthPercent = alpha;
            return true;
        }
        if (source is A.GradientFill gradient && PptxGradientFillCodec.TryRead(gradient, out var semantic))
        {
            fill = new PresentationTableCellFill { GradientFill = semantic };
            return true;
        }
        if (source is A.BlipFill blipFill &&
            PptxImagePaintCodec.TryRead(blipFill, context, out var imagePaint))
        {
            fill = new PresentationTableCellFill { ImagePaint = imagePaint };
            return true;
        }
        return false;
    }

    private static bool TryReadCellBorder(
        OpenXmlElement source,
        out SpreadsheetChartLineStyleArtifact? border)
    {
        border = null;
        var attributes = source.GetAttributes();
        if (attributes.Any(attribute => attribute.LocalName is not ("w" or "cap"))) return false;
        var children = source.ChildElements;
        if (children.Any(child => child is not A.SolidFill and not A.PresetDash and not A.Round and not A.LineJoinBevel and not A.Miter) ||
            children.Count(child => child is A.SolidFill) > 1 ||
            children.Count(child => child is A.PresetDash) > 1 ||
            children.Count(child => child is A.Round or A.LineJoinBevel or A.Miter) > 1)
            return false;
        var solid = source.GetFirstChild<A.SolidFill>();
        if (solid is null || !PptxColor.TryDirectSolidRgbWithOpacity(solid, out var rgb, out var opacity)) return false;
        var dashValue = source.GetFirstChild<A.PresetDash>()?.Val?.Value;
        var dashStyle = dashValue is null ? SpreadsheetChartLineDashStyle.Solid :
            dashValue.Equals(A.PresetLineDashValues.Solid) ? SpreadsheetChartLineDashStyle.Solid :
            dashValue.Equals(A.PresetLineDashValues.Dash) ? SpreadsheetChartLineDashStyle.Dashed :
            dashValue.Equals(A.PresetLineDashValues.Dot) ? SpreadsheetChartLineDashStyle.Dotted :
            dashValue.Equals(A.PresetLineDashValues.DashDot) ? SpreadsheetChartLineDashStyle.DashDot :
            dashValue.Equals(A.PresetLineDashValues.LargeDashDotDot) ? SpreadsheetChartLineDashStyle.DashDotDot :
            SpreadsheetChartLineDashStyle.Unspecified;
        var cap = attributes.FirstOrDefault(attribute =>
            attribute.LocalName == "cap" && attribute.NamespaceUri.Length == 0).Value;
        var output = new SpreadsheetChartLineStyleArtifact
        {
            Color = new SpreadsheetColor { Rgb = rgb },
            DashStyle = dashStyle,
            Cap = cap switch
            {
                null or "flat" => "flat",
                "rnd" => "round",
                "sq" => "square",
                _ => string.Empty,
            },
            Join = children.SingleOrDefault(child => child is A.Round or A.LineJoinBevel or A.Miter)?.LocalName switch
            {
                "round" => "round",
                "bevel" => "bevel",
                "miter" => "miter",
                null => string.Empty,
                _ => string.Empty,
            },
        };
        if (output.DashStyle == SpreadsheetChartLineDashStyle.Unspecified || output.Cap.Length == 0) return false;
        var width = attributes.FirstOrDefault(attribute =>
            attribute.LocalName == "w" && attribute.NamespaceUri.Length == 0).Value;
        if (width is not null)
        {
            if (!long.TryParse(width, NumberStyles.None, CultureInfo.InvariantCulture, out var emu) || emu < 0)
                return false;
            output.WidthPoints = emu / 12_700d;
        }
        if (opacity is { } alpha) output.OpacityThousandthPercent = alpha;
        border = output;
        return true;
    }

    private static void PatchCellProperties(
        A.TableCell cell,
        PresentationTableCell requested,
        PptxPartContext? slideContext)
    {
        var properties = cell.GetFirstChild<A.TableCellProperties>();
        if (properties is null)
        {
            if (requested.Fill is null && requested.Borders is null) return;
            properties = new A.TableCellProperties();
            cell.Append(properties);
        }

        ReplaceCellPaint(properties, requested.Fill, slideContext);
        ReplaceCellBorder(properties, "lnL", requested.Borders?.Left, () => new A.LeftBorderLineProperties());
        ReplaceCellBorder(properties, "lnT", requested.Borders?.Top, () => new A.TopBorderLineProperties());
        ReplaceCellBorder(properties, "lnR", requested.Borders?.Right, () => new A.RightBorderLineProperties());
        ReplaceCellBorder(properties, "lnB", requested.Borders?.Bottom, () => new A.BottomBorderLineProperties());
    }

    private static void PatchCellTextStyle(
        A.TableCell cell,
        PresentationTableCell requested,
        PptxPartContext? slideContext)
    {
        if (requested.TextBody is not null)
        {
            if (requested.TextStyle is not null)
                throw new CodecException("unsupported_presentation_edit", "Source-preserving mixed-run table-cell text must use per-run styles inside text_body, not a uniform textStyle.");
            return;
        }
        if (requested.TextStyle is null && cell.GetFirstChild<A.TextBody>() is null) return;
        var sourceBody = cell.GetFirstChild<A.TextBody>();
        if (sourceBody is null)
        {
            if (requested.TextStyle is not null)
                throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot add a table-cell text body.");
            return;
        }
        var paragraphs = sourceBody.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length < 1 || paragraphs.Any(paragraph =>
            paragraph.ChildElements.Count(child => child is A.Run or A.Field or A.Break) < 1))
        {
            if (requested.TextStyle is not null)
                throw new CodecException("unsupported_presentation_edit", "Source-preserving table-cell text style requires one or more paragraphs with one or more text runs each.");
            return;
        }

        var semantic = PptxTextCodec.ReadDrawingTextBody(sourceBody, slideContext);
        var sourceRuns = paragraphs.SelectMany(paragraph =>
            paragraph.ChildElements.Where(child => child is A.Run or A.Field or A.Break)).ToArray();
        var semanticRuns = semantic.Paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
        if (semantic.Paragraphs.Count != paragraphs.Length || semanticRuns.Length != sourceRuns.Length ||
            semanticRuns.Any(run =>
                run.ContentCase is not (PresentationTextRun.ContentOneofCase.Text or PresentationTextRun.ContentOneofCase.Field or PresentationTextRun.ContentOneofCase.LineBreak) ||
                !CellTextStyleSupported(run)))
            throw new CodecException("unsupported_presentation_edit", "Source-preserving table-cell text style requires paragraphs of direct text, field, or line-break inlines.");
        for (var index = 0; index < semanticRuns.Length; index++)
        {
            var run = semanticRuns[index];
            var replacement = run.ContentCase switch
            {
                PresentationTextRun.ContentOneofCase.Field => new PresentationTextRun { Field = run.Field.Clone() },
                PresentationTextRun.ContentOneofCase.LineBreak => new PresentationTextRun { LineBreak = true },
                _ => new PresentationTextRun { Text = run.Text },
            };
            if (requested.TextStyle is { } style)
            {
                if (style.HasBold) replacement.Bold = style.Bold;
                if (style.HasItalic) replacement.Italic = style.Italic;
                if (style.HasFontSizePoints) replacement.FontSizePoints = style.FontSizePoints;
                if (style.HasFontFamily) replacement.FontFamily = style.FontFamily;
                if (style.HasFontFamilyEastAsia) replacement.FontFamilyEastAsia = style.FontFamilyEastAsia;
                if (style.HasFontFamilyComplexScript) replacement.FontFamilyComplexScript = style.FontFamilyComplexScript;
                if (style.HasColorRgb) replacement.ColorRgb = style.ColorRgb;
                else if (style.HasColorScheme) replacement.ColorScheme = style.ColorScheme;
                if (style.HasColorOpacityThousandthPercent) replacement.ColorOpacityThousandthPercent = style.ColorOpacityThousandthPercent;
                if (style.HasUnderline) replacement.Underline = style.Underline;
                if (style.HasStrike) replacement.Strike = style.Strike;
            }
            var runOffset = 0;
            foreach (var paragraph in semantic.Paragraphs)
            {
                if (index < runOffset + paragraph.Runs.Count)
                {
                    paragraph.Runs[index - runOffset] = replacement;
                    break;
                }
                runOffset += paragraph.Runs.Count;
            }
        }
        var temporary = new P.Shape { TextBody = new P.TextBody { InnerXml = sourceBody.InnerXml } };
        var requestedShape = new PresentationShape
        {
            Text = PptxTextCodec.Flatten(semantic),
            TextBody = semantic,
        };
        PptxTextCodec.Apply(temporary, requestedShape, slideContext!);
        sourceBody.InnerXml = temporary.TextBody!.InnerXml;
    }

    private static void PatchCellTextBody(
        A.TableCell cell,
        PresentationTableCell requested,
        PptxPartContext? slideContext)
    {
        if (requested.TextBody is null) return;
        var sourceBody = cell.GetFirstChild<A.TextBody>();
        if (sourceBody is null)
            throw new CodecException("unsupported_presentation_edit", "Source-preserving PPTX export cannot add a table-cell text body.");
        if (!IsBoundedMixedRunTextBody(requested.TextBody))
            throw new CodecException("unsupported_presentation_edit", "Source-preserving mixed-run table-cell text requires fixed-topology direct text or field runs with bounded direct styles.");
        var sourceSemantic = PptxTextCodec.ReadDrawingTextBody(sourceBody, slideContext);
        ValidateFieldIdentity(sourceSemantic, requested.TextBody);
        var temporary = new P.Shape { TextBody = new P.TextBody { InnerXml = sourceBody.InnerXml } };
        var requestedShape = new PresentationShape
        {
            Text = PptxTextCodec.Flatten(requested.TextBody),
            TextBody = requested.TextBody,
        };
        PptxTextCodec.Apply(temporary, requestedShape, slideContext!);
        sourceBody.InnerXml = temporary.TextBody!.InnerXml;
    }

    private static void ValidateFieldIdentity(PresentationTextBody source, PresentationTextBody requested)
    {
        if (source.Paragraphs.Count != requested.Paragraphs.Count)
            throw new CodecException("presentation_text_topology_changed", "Source-preserving table-cell text requires the original paragraph topology.");
        for (var paragraphIndex = 0; paragraphIndex < source.Paragraphs.Count; paragraphIndex++)
        {
            var sourceRuns = source.Paragraphs[paragraphIndex].Runs;
            var requestedRuns = requested.Paragraphs[paragraphIndex].Runs;
            if (sourceRuns.Count != requestedRuns.Count)
                throw new CodecException("presentation_text_topology_changed", "Source-preserving table-cell text requires the original inline topology.");
            for (var runIndex = 0; runIndex < sourceRuns.Count; runIndex++)
            {
                var sourceRun = sourceRuns[runIndex];
                var requestedRun = requestedRuns[runIndex];
                if (sourceRun.ContentCase != PresentationTextRun.ContentOneofCase.Field ||
                    requestedRun.ContentCase != PresentationTextRun.ContentOneofCase.Field)
                    continue;
                if (!string.Equals(sourceRun.Field.Id, requestedRun.Field.Id, StringComparison.Ordinal) ||
                    !string.Equals(sourceRun.Field.Type, requestedRun.Field.Type, StringComparison.Ordinal))
                    throw new CodecException("unsupported_presentation_edit", "Source-preserving table-cell field edits may change cached text only; field ID and type are source-owned.");
            }
        }
    }

    private static void ReplaceCellPaint(
        A.TableCellProperties properties,
        PresentationTableCellFill? fill,
        PptxPartContext? slideContext)
    {
        var existing = properties.ChildElements.FirstOrDefault(IsCellFill);
        var previousRelationshipId = existing is A.BlipFill previousImage
            ? PptxImagePaintCodec.RelationshipId(previousImage)
            : string.Empty;
        if (fill is null)
        {
            existing?.Remove();
            slideContext?.RemoveIfUnreferenced(previousRelationshipId);
            return;
        }
        var replacement = fill.KindCase switch
        {
            PresentationTableCellFill.KindOneofCase.NoFill => new A.NoFill() as OpenXmlElement,
            PresentationTableCellFill.KindOneofCase.SolidRgb => SolidFill(fill.SolidRgb, fill.HasOpacityThousandthPercent ? fill.OpacityThousandthPercent : null),
            PresentationTableCellFill.KindOneofCase.GradientFill => PptxGradientFillCodec.Build(fill.GradientFill, "Presentation table-cell fill"),
            PresentationTableCellFill.KindOneofCase.ImagePaint => slideContext is null
                ? throw new CodecException("unsupported_presentation_edit", "Source-bound table cell image fills require a containing PresentationML relationship owner.")
                : PptxImagePaintCodec.Build(fill.ImagePaint, slideContext, "Presentation table-cell fill"),
            _ => throw new CodecException("unsupported_presentation_edit", "Source-bound table cell fill is outside the bounded paint profile."),
        };
        if (existing is null) properties.InsertAt(replacement, 0);
        else properties.ReplaceChild(replacement, existing);
        slideContext?.RemoveIfUnreferenced(previousRelationshipId);
    }

    private static void ReplaceCellBorder(
        A.TableCellProperties properties,
        string localName,
        SpreadsheetChartLineStyleArtifact? border,
        Func<OpenXmlElement> create)
    {
        var existing = properties.ChildElements.FirstOrDefault(child => child.LocalName == localName && child.NamespaceUri == DrawingNamespace);
        if (border is null)
        {
            existing?.Remove();
            return;
        }
        var replacement = create();
        ApplyBorder(replacement, border);
        if (existing is null) properties.Append(replacement);
        else properties.ReplaceChild(replacement, existing);
    }

    private static void ApplyBorder(OpenXmlElement output, SpreadsheetChartLineStyleArtifact source)
    {
        if (source.HasWidthPoints)
            output.SetAttribute(new OpenXmlAttribute("w", string.Empty,
                checked((long)Math.Round(source.WidthPoints * 12_700, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture)));
        if (source.Cap.Length > 0)
            output.SetAttribute(new OpenXmlAttribute("cap", string.Empty, source.Cap switch
            {
                "round" => "rnd",
                "square" => "sq",
                _ => "flat",
            }));
        if (source.Color is { SourceCase: SpreadsheetColor.SourceOneofCase.Rgb } color)
            output.Append(SolidFill(color.Rgb, source.HasOpacityThousandthPercent ? source.OpacityThousandthPercent : null));
        if (source.DashStyle != SpreadsheetChartLineDashStyle.Unspecified)
            output.Append(new A.PresetDash { Val = source.DashStyle switch
            {
                SpreadsheetChartLineDashStyle.Dashed => A.PresetLineDashValues.Dash,
                SpreadsheetChartLineDashStyle.Dotted => A.PresetLineDashValues.Dot,
                SpreadsheetChartLineDashStyle.DashDot => A.PresetLineDashValues.DashDot,
                SpreadsheetChartLineDashStyle.DashDotDot => A.PresetLineDashValues.LargeDashDotDot,
                _ => A.PresetLineDashValues.Solid,
            }});
        if (source.Join.Length > 0)
            output.Append(source.Join switch
            {
                "round" => new A.Round(),
                "bevel" => new A.LineJoinBevel(),
                _ => new A.Miter(),
            });
    }

    private static bool TryReadMergeRanges(IReadOnlyList<A.TableCell[]> nativeRows, PresentationTable table)
    {
        var rowCount = nativeRows.Count;
        var columnCount = nativeRows[0].Length;
        var cells = new NativeMergeCell[rowCount, columnCount];
        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cell = nativeRows[row][column];
                var hasRowSpan = cell.RowSpan is not null;
                var hasColumnSpan = cell.GridSpan is not null;
                var hasHorizontal = cell.HorizontalMerge is not null;
                var hasVertical = cell.VerticalMerge is not null;
                var rowSpan = cell.RowSpan?.Value ?? 1;
                var columnSpan = cell.GridSpan?.Value ?? 1;
                var horizontal = cell.HorizontalMerge?.Value ?? false;
                var vertical = cell.VerticalMerge?.Value ?? false;
                if ((hasRowSpan && rowSpan <= 1) || (hasColumnSpan && columnSpan <= 1) ||
                    (hasHorizontal && !horizontal) || (hasVertical && !vertical) ||
                    ((horizontal || vertical) && (hasRowSpan || hasColumnSpan)))
                    return false;
                cells[row, column] = new NativeMergeCell(rowSpan, columnSpan, horizontal, vertical, hasRowSpan, hasColumnSpan, hasHorizontal, hasVertical);
            }
        }

        var expected = new Dictionary<(int Row, int Column), MergeCellPlan>();
        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cell = cells[row, column];
                if (cell.HorizontalMerge || cell.VerticalMerge || cell.RowSpan == 1 && cell.ColumnSpan == 1) continue;
                if ((long)row + cell.RowSpan > rowCount || (long)column + cell.ColumnSpan > columnCount) return false;
                var range = new PresentationTableMergeRange
                {
                    StartRow = (uint)row,
                    EndRow = (uint)(row + cell.RowSpan - 1),
                    StartColumn = (uint)column,
                    EndColumn = (uint)(column + cell.ColumnSpan - 1),
                };
                for (var coveredRow = row; coveredRow <= range.EndRow; coveredRow++)
                {
                    for (var coveredColumn = column; coveredColumn <= range.EndColumn; coveredColumn++)
                    {
                        var isOrigin = coveredRow == row && coveredColumn == column;
                        if (!expected.TryAdd((coveredRow, coveredColumn), new MergeCellPlan(
                            isOrigin,
                            isOrigin ? cell.RowSpan : 0,
                            isOrigin ? cell.ColumnSpan : 0,
                            !isOrigin && coveredColumn > column,
                            !isOrigin && coveredRow > row)))
                            return false;
                    }
                }
                table.MergeRanges.Add(range);
            }
        }

        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cell = cells[row, column];
                if (expected.TryGetValue((row, column), out var planned))
                {
                    if (planned.IsOrigin)
                    {
                        if (cell.HorizontalMerge || cell.VerticalMerge || cell.RowSpan != planned.RowSpan || cell.ColumnSpan != planned.ColumnSpan) return false;
                    }
                    else
                    {
                        if (cell.HasRowSpan || cell.HasColumnSpan || cell.HorizontalMerge != planned.HorizontalMerge || cell.VerticalMerge != planned.VerticalMerge ||
                            !string.IsNullOrEmpty(table.Rows[row].Cells[column].Text))
                            return false;
                    }
                }
                else if (cell.HasRowSpan || cell.HasColumnSpan || cell.HasHorizontalMerge || cell.HasVerticalMerge)
                {
                    return false;
                }
            }
        }
        return true;
    }

    private static Dictionary<(int Row, int Column), MergeCellPlan> CreateMergePlan(PresentationTable table, string elementId)
    {
        var plan = new Dictionary<(int Row, int Column), MergeCellPlan>();
        for (var rangeIndex = 0; rangeIndex < table.MergeRanges.Count; rangeIndex++)
        {
            var range = table.MergeRanges[rangeIndex];
            if (range.EndRow < range.StartRow || range.EndColumn < range.StartColumn ||
                range.EndRow >= table.Rows.Count || range.EndColumn >= table.ColumnWidthsEmu.Count ||
                range.StartRow == range.EndRow && range.StartColumn == range.EndColumn)
                throw Invalid(elementId, $"merge range {rangeIndex} must cover at least two in-bounds cells");
            var rowSpan = checked((int)(range.EndRow - range.StartRow + 1));
            var columnSpan = checked((int)(range.EndColumn - range.StartColumn + 1));
            for (var row = checked((int)range.StartRow); row <= range.EndRow; row++)
            {
                for (var column = checked((int)range.StartColumn); column <= range.EndColumn; column++)
                {
                    var isOrigin = row == range.StartRow && column == range.StartColumn;
                    if (!plan.TryAdd((row, column), new MergeCellPlan(
                        isOrigin,
                        isOrigin ? rowSpan : 0,
                        isOrigin ? columnSpan : 0,
                        !isOrigin && column > range.StartColumn,
                        !isOrigin && row > range.StartRow)))
                        throw Invalid(elementId, $"merge ranges overlap at cell {row},{column}");
                    if (!isOrigin && !string.IsNullOrEmpty(table.Rows[row].Cells[column].Text))
                        throw Invalid(elementId, $"covered merge cell {row},{column} must be empty");
                }
            }
        }
        return plan;
    }

    private static A.TableCell BuildCell(
        PresentationTableCell source,
        bool header,
        PresentationTable table,
        PptxPartContext slideContext)
    {
        if (source.TextBody is not null)
            return new A.TableCell(
                PptxTextCodec.BuildDrawingTextBody(source.TextBody, slideContext),
                BuildCellProperties(source, header, table, slideContext));

        var style = table.DefaultTextStyle;
        var runProperties = new A.RunProperties
        {
            Language = "en-US",
            FontSize = style?.HasFontSizePoints == true
                ? checked((int)Math.Round(style.FontSizePoints * 100))
                : 1_350,
            Bold = style?.HasBold == true ? style.Bold || header : header,
            Italic = style?.HasItalic == true ? style.Italic : null,
        };
        var colorOpacity = style?.HasColorOpacityThousandthPercent == true
            ? style.ColorOpacityThousandthPercent
            : (uint?)null;
        runProperties.Append(style?.HasColorScheme == true
            ? PptxColor.BuildSolidScheme(style.ColorScheme, colorOpacity)
            : PptxColor.BuildSolidRgb(style?.HasColorRgb == true ? style.ColorRgb : header ? "000000" : "0F172A", colorOpacity));
        if (style?.HasFontFamily == true)
        {
            runProperties.Append(new A.LatinFont { Typeface = style.FontFamily });
            runProperties.Append(new A.EastAsianFont { Typeface = style.HasFontFamilyEastAsia ? style.FontFamilyEastAsia : style.FontFamily });
        }
        if (style?.HasFontFamilyComplexScript == true)
            runProperties.Append(new A.ComplexScriptFont { Typeface = style.FontFamilyComplexScript });
        var paragraphs = source.Text.Split('\n').Select(line => new A.Paragraph(
            new A.Run((A.RunProperties)runProperties.CloneNode(true), new A.Text(line)),
            new A.EndParagraphRunProperties { Language = "en-US", FontSize = 1_350 })).ToArray();
        var textBody = new A.TextBody(new A.BodyProperties(), new A.ListStyle());
        textBody.Append(paragraphs);
        return new A.TableCell(
            textBody,
            BuildCellProperties(source, header, table, slideContext));
    }

    private static A.TableCellProperties BuildCellProperties(
        PresentationTableCell source,
        bool header,
        PresentationTable table,
        PptxPartContext slideContext)
    {
        var output = new A.TableCellProperties();
        if (source.Borders is { } borders)
        {
            if (borders.Left is not null) output.Append(BuildBorder<A.LeftBorderLineProperties>(borders.Left));
            if (borders.Right is not null) output.Append(BuildBorder<A.RightBorderLineProperties>(borders.Right));
            if (borders.Top is not null) output.Append(BuildBorder<A.TopBorderLineProperties>(borders.Top));
            if (borders.Bottom is not null) output.Append(BuildBorder<A.BottomBorderLineProperties>(borders.Bottom));
        }
        if (source.Fill is { } fill)
        {
            output.Append(fill.KindCase switch
            {
                PresentationTableCellFill.KindOneofCase.NoFill => new A.NoFill(),
                PresentationTableCellFill.KindOneofCase.SolidRgb => SolidFill(fill.SolidRgb, fill.HasOpacityThousandthPercent ? fill.OpacityThousandthPercent : null),
                PresentationTableCellFill.KindOneofCase.GradientFill => PptxGradientFillCodec.Build(fill.GradientFill, "Presentation table-cell fill"),
                PresentationTableCellFill.KindOneofCase.ImagePaint => PptxImagePaintCodec.Build(fill.ImagePaint, slideContext, "Presentation table-cell fill"),
                _ => throw new InvalidOperationException("Validated presentation table-cell fill changed unexpectedly."),
            });
        }
        else if (table.DefaultCellFillCase == PresentationTable.DefaultCellFillOneofCase.NoDefaultCellFill)
            output.Append(new A.NoFill());
        else output.Append(SolidFill(
            table.DefaultCellFillCase == PresentationTable.DefaultCellFillOneofCase.DefaultCellFillRgb
                ? table.DefaultCellFillRgb
                : header ? "EDEDED" : "FFFFFF",
            null));
        return output;
    }

    private static A.SolidFill SolidFill(string rgb, uint? opacity)
    {
        var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(rgb) };
        if (opacity is { } alpha) color.Append(new A.Alpha { Val = checked((int)alpha) });
        return new A.SolidFill(color);
    }

    private static T BuildBorder<T>(SpreadsheetChartLineStyleArtifact source)
        where T : OpenXmlCompositeElement, new()
    {
        var output = new T();
        if (source.HasWidthPoints)
            output.SetAttribute(new OpenXmlAttribute("w", string.Empty, checked((long)Math.Round(source.WidthPoints * 12_700, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture)));
        if (source.Cap.Length > 0)
            output.SetAttribute(new OpenXmlAttribute("cap", string.Empty, source.Cap switch
            {
                "round" => "rnd",
                "square" => "sq",
                _ => "flat",
            }));
        if (source.Color is not null)
            output.Append(SolidFill(source.Color.Rgb, source.HasOpacityThousandthPercent ? source.OpacityThousandthPercent : null));
        if (source.DashStyle != SpreadsheetChartLineDashStyle.Unspecified)
            output.Append(new A.PresetDash { Val = source.DashStyle switch
            {
                SpreadsheetChartLineDashStyle.Dashed => A.PresetLineDashValues.Dash,
                SpreadsheetChartLineDashStyle.Dotted => A.PresetLineDashValues.Dot,
                SpreadsheetChartLineDashStyle.DashDot => A.PresetLineDashValues.DashDot,
                SpreadsheetChartLineDashStyle.DashDotDot => A.PresetLineDashValues.LargeDashDotDot,
                _ => A.PresetLineDashValues.Solid,
            }});
        if (source.Join.Length > 0) output.Append(source.Join switch
        {
            "round" => new A.Round(),
            "bevel" => new A.LineJoinBevel(),
            _ => new A.Miter(),
        });
        return output;
    }

    private static void ValidateCellFill(
        PresentationTableCellFill? fill,
        string elementId,
        PptxAssetCatalog? assets)
    {
        if (fill is null) return;
        if (fill.KindCase == PresentationTableCellFill.KindOneofCase.None)
            throw Invalid(elementId, "cell fill must select none, solid, gradient, or image");
        if (fill.KindCase == PresentationTableCellFill.KindOneofCase.NoFill && !fill.NoFill)
            throw Invalid(elementId, "cell no_fill must be true when selected");
        if (fill.KindCase == PresentationTableCellFill.KindOneofCase.SolidRgb)
            PptxColor.Normalize(fill.SolidRgb);
        if (fill.KindCase == PresentationTableCellFill.KindOneofCase.GradientFill)
            PptxGradientFillCodec.Validate(fill.GradientFill, $"Presentation table {elementId} cell fill");
        if (fill.KindCase == PresentationTableCellFill.KindOneofCase.ImagePaint)
            PptxImagePaintCodec.Validate(fill.ImagePaint, $"Presentation table {elementId} cell fill", assets);
        if (fill.HasOpacityThousandthPercent &&
            (fill.KindCase != PresentationTableCellFill.KindOneofCase.SolidRgb || fill.OpacityThousandthPercent > 100_000))
            throw Invalid(elementId, "cell fill opacity requires a solid RGB fill and must be 0 through 100000");
    }

    private static void ValidateCellBorders(PresentationTableCellBorders? borders, string elementId)
    {
        if (borders is null) return;
        XlsxChartSeriesLineStyleCodec.ValidateLine(borders.Left, "presentation", elementId, "cell", "left border");
        XlsxChartSeriesLineStyleCodec.ValidateLine(borders.Top, "presentation", elementId, "cell", "top border");
        XlsxChartSeriesLineStyleCodec.ValidateLine(borders.Right, "presentation", elementId, "cell", "right border");
        XlsxChartSeriesLineStyleCodec.ValidateLine(borders.Bottom, "presentation", elementId, "cell", "bottom border");
    }

    private static IReadOnlyList<A.Text>? TextLeaves(A.TableCell cell)
    {
        if (cell.ChildElements.Count == 0) return null;
        var body = cell.GetFirstChild<A.TextBody>();
        if (body is null) return null;
        var paragraphs = body.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length == 0 || paragraphs.Any(paragraph => paragraph.Elements<A.Run>().Count() != 1 || paragraph.Descendants<A.Text>().Count() != 1)) return null;
        var leaves = new List<A.Text>(paragraphs.Length);
        foreach (var paragraph in paragraphs)
        {
            var runs = paragraph.Elements<A.Run>().ToArray();
            var textLeaves = paragraph.Descendants<A.Text>().ToArray();
            if (runs.Length != 1 || textLeaves.Length != 1 ||
                textLeaves.Any(value => value.Parent is not A.Run || value.Text.Contains('\n'))) return null;
            leaves.Add(textLeaves[0]);
        }
        return leaves;
    }

    private static IReadOnlyList<IReadOnlyList<A.Text>>? TextRunLeaves(A.TableCell cell)
    {
        if (cell.ChildElements.Count == 0) return null;
        var body = cell.GetFirstChild<A.TextBody>();
        if (body is null) return null;
        var paragraphs = body.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length == 0) return null;
        var output = new List<IReadOnlyList<A.Text>>(paragraphs.Length);
        foreach (var paragraph in paragraphs)
        {
            var runs = paragraph.Elements<A.Run>().ToArray();
            if (runs.Length == 0) return null;
            var leaves = new List<A.Text>(runs.Length);
            foreach (var run in runs)
            {
                if (run.ChildElements.Any(child => child is not A.RunProperties and not A.Text)) return null;
                var text = run.GetFirstChild<A.Text>();
                if (text is null || run.Elements<A.Text>().Count() != 1 || text.Text?.Contains('\n') == true) return null;
                leaves.Add(text);
            }
            output.Add(leaves);
        }
        return output;
    }

    private static void SetParagraphText(IReadOnlyList<A.Text> leaves, string value)
    {
        if (leaves.Count == 1)
        {
            leaves[0].Text = value;
            return;
        }
        var sourceLengths = leaves.Select(leaf => leaf.Text?.Length ?? 0).ToArray();
        var sourceTotal = sourceLengths.Sum();
        var offset = 0;
        for (var index = 0; index < leaves.Count; index++)
        {
            var length = index == leaves.Count - 1
                ? value.Length - offset
                : sourceTotal <= 0
                    ? 0
                    : (int)Math.Round(value.Length * (double)sourceLengths[index] / sourceTotal, MidpointRounding.ToEven);
            length = Math.Clamp(length, 0, value.Length - offset);
            leaves[index].Text = value.Substring(offset, length);
            offset += length;
        }
    }

    private static string CellText(A.TableCell cell)
    {
        var body = cell.GetFirstChild<A.TextBody>();
        return body is null
            ? string.Empty
            : string.Join("\n", body.Elements<A.Paragraph>().Select(paragraph =>
                string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text))));
    }

    private static bool TablePropertiesSupported(A.TableProperties properties)
    {
        if (!HasOnlyAttributes(properties, "firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol")) return false;
        var styleIds = properties.Elements<A.TableStyleId>().ToArray();
        var noFills = properties.Elements<A.NoFill>().ToArray();
        return properties.ChildElements.All(child => child is A.TableStyleId or A.NoFill) &&
               styleIds.Length <= 1 && noFills.Length <= 1 &&
               styleIds.All(style => !style.HasAttributes && !style.HasChildren && style.InnerText.Length <= 256) &&
               noFills.All(fill => !fill.HasAttributes && !fill.HasChildren);
    }

    private static bool GridColumnSupported(A.GridColumn column) =>
        HasOnlyAttributes(column, "w") &&
        column.ChildElements.All(child => child is A.ExtensionList && ExtensionListSupported((A.ExtensionList)child));

    private static bool TableRowSupported(A.TableRow row) =>
        HasOnlyAttributes(row, "h") &&
        row.ChildElements.All(child => child is A.TableCell || child is A.ExtensionList && ExtensionListSupported((A.ExtensionList)child));

    private static bool ExtensionListSupported(A.ExtensionList extensions)
    {
        if (extensions.HasAttributes) return false;
        foreach (var extension in extensions.ChildElements)
        {
            if (extension is not A.Extension || !HasOnlyAttributes(extension, "uri") || extension.ChildElements.Count != 1)
                return false;
            if (string.IsNullOrWhiteSpace(extension.GetAttributes().Single().Value)) return false;
        }
        return true;
    }

    private static bool ScaledExtentSupported(long frameExtent, long contentExtent)
    {
        if (frameExtent <= 0 || contentExtent <= 0) return false;
        // Permit ordinary graphic-frame scaling while bounding malformed
        // source dimensions that could otherwise produce extreme geometry.
        return (double)frameExtent / contentExtent is >= 1d / 256d and <= 256d;
    }

    private static void SetFrame(P.Transform transform, PresentationTable table)
    {
        transform.Offset!.X = table.LeftEmu;
        transform.Offset.Y = table.TopEmu;
        transform.Extents!.Cx = table.WidthEmu;
        transform.Extents.Cy = table.HeightEmu;
        PptxFrameTransformCodec.Apply(transform, table.FrameTransform);
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }

    private static long Sum(IEnumerable<long> values, string elementId)
    {
        try { return values.Aggregate(0L, checked((total, value) => total + value)); }
        catch (OverflowException) { throw Invalid(elementId, "grid dimensions overflow the supported EMU range"); }
    }

    private static CodecException Invalid(string elementId, string message) =>
        new("invalid_presentation_table", $"Presentation table {elementId} {message}.");
}
