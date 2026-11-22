using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Rendering;

namespace Aid.Xlsx.Model.Grid.Columns;

internal class XlsxColumn : XlsxRangeBase, IXlsxColumn, IXlsxFormatContainer
{
    private int _outlineLevel;

    /// <summary>
    /// The direct constructor should only be used in <see cref="XlsxWorksheet.RangeFactory"/>.
    /// </summary>
    public XlsxColumn(XlsxWorksheet worksheet, Int32 column)
        : base(XlsxRangeAddress.EntireColumn(worksheet, column))
    {
        SetColumnNumber(column);

        Width = worksheet.ColumnWidth;
    }

    /// <summary>
    /// Get area of this column.
    /// </summary>
    internal XlsxColumnArea Area => new(Worksheet.Name, ColumnNumber());

    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.Column; }
    }

    public Boolean Collapsed { get; set; }

    #region IXlsxColumn Members

    public Double Width { get; set; }

    IXlsxCells IXlsxColumn.Cells(String cellsInColumn) => Cells(cellsInColumn);

    IXlsxCells IXlsxColumn.Cells(Int32 firstRow, Int32 lastRow) => Cells(firstRow, lastRow);

    public void Delete()
    {
        int columnNumber = ColumnNumber();
        Delete(XlsxShiftDeletedCells.ShiftCellsLeft);
        Worksheet.DeleteColumn(columnNumber);
    }

    public new IXlsxColumn Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXlsxCell Cell(Int32 rowNumber)
    {
        return Cell(rowNumber, 1);
    }

    public override XlsxCells Cells(String cellsInColumn)
    {
        var retVal = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.All);
        var rangePairs = cellsInColumn.Split(',');
        foreach (string pair in rangePairs)
            retVal.Add(Range(pair.Trim()).RangeAddress);
        return retVal;
    }

    public override IXlsxCells Cells()
    {
        return Cells(true, XlsxCellsUsedOptions.All);
    }

    public override XlsxCells Cells(Boolean usedCellsOnly)
    {
        if (usedCellsOnly)
            return Cells(true, XlsxCellsUsedOptions.AllContents);
        else
            return Cells(FirstCellUsed().Address.RowNumber, LastCellUsed().Address.RowNumber);
    }

    public XlsxCells Cells(Int32 firstRow, Int32 lastRow)
    {
        return Cells(firstRow + ":" + lastRow);
    }

    public new IXlsxColumns InsertColumnsAfter(Int32 numberOfColumns)
    {
        int columnNum = ColumnNumber();
        Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
        Worksheet.Column(columnNum).InsertColumnsAfterVoid(true, numberOfColumns);
        var newColumns = Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
        foreach (var newColumn in newColumns)
        {
            var internalColumn = Worksheet.Internals.ColumnsCollection[newColumn.ColumnNumber()];
            internalColumn.Width = Width;
            internalColumn.FormatValue = FormatValue; // Is within a worbook
            internalColumn.Collapsed = Collapsed;
            internalColumn.IsHidden = IsHidden;
            internalColumn._outlineLevel = OutlineLevel;
        }

        return newColumns;
    }

    public new IXlsxColumns InsertColumnsBefore(Int32 numberOfColumns)
    {
        int columnNum = ColumnNumber();
        if (columnNum > 1)
        {
            return Worksheet.Column(columnNum - 1).InsertColumnsAfter(numberOfColumns);
        }

        Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);
        Worksheet.Column(columnNum).InsertColumnsBeforeVoid(true, numberOfColumns);

        return Worksheet.Columns(columnNum, columnNum + numberOfColumns - 1);
    }

    public IXlsxColumn AdjustToContents()
    {
        return AdjustToContents(1);
    }

    public IXlsxColumn AdjustToContents(Int32 startRow)
    {
        return AdjustToContents(startRow, XlsxHelper.MaxRowNumber);
    }

    public IXlsxColumn AdjustToContents(Int32 startRow, Int32 endRow)
    {
        return AdjustToContents(startRow, endRow, 0, Double.MaxValue);
    }

    public IXlsxColumn AdjustToContents(Double minWidth, Double maxWidth)
    {
        return AdjustToContents(1, XlsxHelper.MaxRowNumber, minWidth, maxWidth);
    }

    public IXlsxColumn AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
    {
        return AdjustToContents(startRow, XlsxHelper.MaxRowNumber, minWidth, maxWidth);
    }

    public IXlsxColumn AdjustToContents(Int32 startRow, Int32 endRow, Double minWidthNoC, Double maxWidthNoC)
    {
        var engine = Worksheet.Workbook.GraphicEngine;
        var dpi = new Dpi(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);
        var columnWidthPx = CalculateMinColumnWidth(startRow, endRow, engine, dpi);

        // Maximum digit width, rounded to pixels, so Calibri at 11 pts returns 7 pixels MDW (the correct value)
        var mdw = (int)Math.Round(engine.GetMaxDigitWidth(Worksheet.Workbook.Format.Font, dpi.X));

        var minWidthInPx = Math.Ceiling(XlsxHelper.NoCToPixels(minWidthNoC, mdw));
        if (columnWidthPx < minWidthInPx)
            columnWidthPx = (int)minWidthInPx;

        var maxWidthInPx = Math.Ceiling(XlsxHelper.NoCToPixels(maxWidthNoC, mdw));
        if (columnWidthPx > maxWidthInPx)
            columnWidthPx = (int)maxWidthInPx;

        var colMaxWidth = XlsxHelper.PixelToNoC(columnWidthPx, mdw);

        // If there is nothing in the column, use worksheet column width.
        if (colMaxWidth <= 0)
            colMaxWidth = Worksheet.ColumnWidth;

        Width = colMaxWidth;

        return this;
    }

    /// <summary>
    /// Calculate column width in pixels according to the content of cells.
    /// </summary>
    /// <param name="startRow">First row number whose content is used for determination.</param>
    /// <param name="endRow">Last row number whose content is used for determination.</param>
    /// <param name="engine">Engine to determine size of glyphs.</param>
    /// <param name="dpi">DPI of the worksheet.</param>
    private int CalculateMinColumnWidth(int startRow, int endRow, IXlsxGraphicEngine engine, Dpi dpi)
    {
        var autoFilterRows = new List<Int32>();
        if (this.Worksheet.AutoFilter != null && Worksheet.AutoFilter.Range != null)
            autoFilterRows.Add(this.Worksheet.AutoFilter.Range.FirstRow().RowNumber());

        autoFilterRows.AddRange(Worksheet.Tables.Where<XlsxTable>(t =>
                t.AutoFilter != null
                && t.AutoFilter.Range != null
                && !autoFilterRows.Contains(t.AutoFilter.Range.FirstRow().RowNumber()))
            .Select(t => t.AutoFilter.Range.FirstRow().RowNumber()));

        // Cache MDW for each font to avoid too many allocations
        var scaledMdwMap = new Dictionary<XlsxFontFormatValue, double>(ReferenceEqualityComparer<XlsxFontFormatValue>.Instance);

        // Reusable buffer
        var glyphs = new List<GlyphBox>();
        var columnWidthPx = 0;
        foreach (var cell in Column(startRow, endRow).CellsUsed())
        {
            // Clear maintains capacity -> reduce need for GC
            glyphs.Clear();

            if (cell.IsMerged())
                continue;

            var cellStyle = Worksheet.GetStyleValue(cell.Point);

            cell.GetGlyphBoxes(engine, dpi, glyphs);
            var textWidthPx = (int)Math.Ceiling(GetContentWidth(cellStyle.Alignment.TextRotation.Value, glyphs));

            if (!scaledMdwMap.TryGetValue(cellStyle.Font, out var scaledMdw))
            {
                var mdw = engine.GetMaxDigitWidth(cellStyle.Font.ToFontBase(), dpi.X);
                scaledMdw = Math.Round(mdw, MidpointRounding.AwayFromZero);
                scaledMdwMap.Add(cellStyle.Font, scaledMdw);
            }

            // Not sure about rounding, but larger is probably better, so use ceiling.
            // Due to mismatched rendering, add 3% instead of 1.75%, to have additional space.
            var oneSidePadding = (int)Math.Ceiling(textWidthPx * 0.03 + scaledMdw / 4);

            // Cell width if calculated as content width + padding on each side of a content.
            // The one side padding is roughly 1.75% of content + MDW/4.
            // The additional pixel is there for lines between cells.
            var cellWidthPx = textWidthPx + 2 * oneSidePadding + 1;

            if (autoFilterRows.Contains(cell.Address.RowNumber))
            {
                // Autofilter arrow is 16px at 96dpi, scaling through DPI, e.g. 20px at 120dpi
                cellWidthPx += (int)Math.Round(16d * dpi.X / 96d, MidpointRounding.AwayFromZero);
            }

            columnWidthPx = Math.Max(cellWidthPx, columnWidthPx);
        }

        return columnWidthPx;
    }

    private static double GetContentWidth(int textRotationDeg, List<GlyphBox> glyphs)
    {
        if (textRotationDeg == 0)
        {
            var maxTextWidth = 0d;
            var lineTextWidth = 0d;
            foreach (var glyph in glyphs)
            {
                if (!glyph.IsLineBreak)
                {
                    lineTextWidth += glyph.AdvanceWidth;
                    maxTextWidth = Math.Max(lineTextWidth, maxTextWidth);
                }
                else
                    lineTextWidth = 0;
            }

            return maxTextWidth;
        }
        if (textRotationDeg == 255)
        {
            // Glyphs are arranged vertically, top to bottom.
            var maxGlyphWidth = 0d;
            foreach (var grapheme in glyphs)
                maxGlyphWidth = Math.Max(grapheme.AdvanceWidth, maxGlyphWidth);

            return maxGlyphWidth;
        }
        else
        {
            // Glyphs are rotated
            if (textRotationDeg > 90)
                textRotationDeg = 90 - textRotationDeg;

            var totalWidth = 0d;
            var maxHeight = 0d;
            foreach (var glyph in glyphs)
            {
                totalWidth += glyph.AdvanceWidth;
                maxHeight = Math.Max(maxHeight, glyph.LineHeight);
            }

            var projectedHeight = maxHeight * Math.Cos(XlsxHelper.DegToRad(90 - textRotationDeg));
            var projectedWidth = totalWidth * Math.Cos(XlsxHelper.DegToRad(textRotationDeg));
            return projectedWidth + projectedHeight;
        }
    }

    public IXlsxColumn Hide()
    {
        IsHidden = true;
        return this;
    }

    public IXlsxColumn Unhide()
    {
        IsHidden = false;
        return this;
    }

    public Boolean IsHidden { get; set; }

    public Int32 OutlineLevel
    {
        get { return _outlineLevel; }
        set
        {
            if (value < 0 || value > 8)
                throw new ArgumentOutOfRangeException("value", "Outline level must be between 0 and 8.");

            Worksheet.IncrementColumnOutline(value);
            Worksheet.DecrementColumnOutline(_outlineLevel);
            _outlineLevel = value;
        }
    }

    public IXlsxColumn Group()
    {
        return Group(false);
    }

    public IXlsxColumn Group(Boolean collapse)
    {
        if (OutlineLevel < 8)
            OutlineLevel += 1;

        Collapsed = collapse;
        return this;
    }

    public IXlsxColumn Group(Int32 outlineLevel)
    {
        return Group(outlineLevel, false);
    }

    public IXlsxColumn Group(Int32 outlineLevel, Boolean collapse)
    {
        OutlineLevel = outlineLevel;
        Collapsed = collapse;
        return this;
    }

    public IXlsxColumn Ungroup()
    {
        return Ungroup(false);
    }

    public IXlsxColumn Ungroup(Boolean ungroupFromAll)
    {
        if (ungroupFromAll)
            OutlineLevel = 0;
        else
        {
            if (OutlineLevel > 0)
                OutlineLevel -= 1;
        }
        return this;
    }

    public IXlsxColumn Collapse()
    {
        Collapsed = true;
        return Hide();
    }

    public IXlsxColumn Expand()
    {
        Collapsed = false;
        return Unhide();
    }

    public Int32 CellCount()
    {
        return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
    }

    public IXlsxColumn Sort(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false,
                          Boolean ignoreBlanks = true)
    {
        Sort(1, sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    IXlsxRangeColumn IXlsxColumn.Column(Int32 start, Int32 end) => Column(start, end);

    IXlsxRangeColumn IXlsxColumn.CopyTo(IXlsxCell target)
    {
        var copy = AsRange().CopyTo(target);
        return copy.Column(1);
    }

    IXlsxRangeColumn IXlsxColumn.CopyTo(IXlsxRangeBase target)
    {
        var copy = AsRange().CopyTo(target);
        return copy.Column(1);
    }

    public IXlsxColumn CopyTo(IXlsxColumn column)
    {
        column.Clear();
        var newColumn = (XlsxColumn)column;
        newColumn.Width = Width;
        newColumn.FormatValue = newColumn.Worksheet.Workbook.Styles.GetRegisteredCellFormat(GetFormat());
        newColumn.IsHidden = IsHidden;

        (this as XlsxRangeBase).CopyTo(column);

        return newColumn;
    }

    public XlsxRangeColumn Column(Int32 start, Int32 end)
    {
        return Range(start, 1, end, 1).Column(1);
    }

    public IXlsxRangeColumn Column(IXlsxCell start, IXlsxCell end)
    {
        return Column(start.Address.RowNumber, end.Address.RowNumber);
    }

    public IXlsxRangeColumns Columns(String columns)
    {
        var retVal = new XlsxRangeColumns(Worksheet);
        var columnPairs = columns.Split(',');
        foreach (string pair in columnPairs)
            AsRange().Columns(pair.Trim()).ForEach(retVal.Add);
        return retVal;
    }

    /// <summary>
    ///   Adds a vertical page break after this column.
    /// </summary>
    public IXlsxColumn AddVerticalPageBreak()
    {
        Worksheet.PageSetup.AddVerticalPageBreak(ColumnNumber());
        return this;
    }

    public IXlsxRangeColumn ColumnUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents)
    {
        return Column((this as IXlsxRangeBase).FirstCellUsed(options),
                      (this as IXlsxRangeBase).LastCellUsed(options));
    }

    #endregion IXlsxColumn Members

    #region IXlsxFormatContainer

    /// <remarks>
    /// Format of the column or <c>null</c> for default format.
    /// </remarks>
    /// <inheritdoc cref="IXlsxFormatContainer.FormatValue"/>
    public XlsxCellFormatValue? FormatValue { get; set; }

    internal override XlsxCellFormat Format => XlsxCellFormat.ForColumn(this);

    private XlsxCellFormatValue GetFormat()
    {
        return FormatValue ?? Worksheet.Workbook.Styles.DefaultCellFormat;
    }

    #endregion

    public override XlsxRange AsRange()
    {
        return Range(1, 1, XlsxHelper.MaxRowNumber, 1);
    }

    internal override void WorksheetRangeShiftedColumns(XlsxRange range, int columnsShifted)
    {
        return; // Columns are shifted by XlsxColumnCollection
    }

    internal override void WorksheetRangeShiftedRows(XlsxRange range, int rowsShifted)
    {
        //do nothing
    }

    internal void SetColumnNumber(int column)
    {
        RangeAddress = new XlsxRangeAddress(
            new XlsxAddress(Worksheet,
                          1,
                          column,
                          RangeAddress.FirstAddress.FixedRow,
                          RangeAddress.FirstAddress.FixedColumn),
            new XlsxAddress(Worksheet,
                          XlsxHelper.MaxRowNumber,
                          column,
                          RangeAddress.LastAddress.FixedRow,
                          RangeAddress.LastAddress.FixedColumn));
    }

    public override XlsxRange Range(String rangeAddressStr)
    {
        String rangeAddressToUse;
        if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
        {
            if (rangeAddressStr.Contains('-'))
                rangeAddressStr = rangeAddressStr.Replace('-', ':');

            var arrRange = rangeAddressStr.Split(':');
            string firstPart = arrRange[0];
            string secondPart = arrRange[1];
            rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
        }
        else
            rangeAddressToUse = FixColumnAddress(rangeAddressStr);

        var rangeAddress = new XlsxRangeAddress(Worksheet, rangeAddressToUse);
        return Range(rangeAddress);
    }

    public IXlsxRangeColumn Range(int firstRow, int lastRow)
    {
        return Range(firstRow, 1, lastRow, 1).Column(1);
    }

    private XlsxColumn ColumnShift(Int32 columnsToShift)
    {
        return Worksheet.Column(ColumnNumber() + columnsToShift);
    }

    #region XlsxColumn Left

    IXlsxColumn IXlsxColumn.ColumnLeft()
    {
        return ColumnLeft();
    }

    IXlsxColumn IXlsxColumn.ColumnLeft(Int32 step)
    {
        return ColumnLeft(step);
    }

    public XlsxColumn ColumnLeft()
    {
        return ColumnLeft(1);
    }

    public XlsxColumn ColumnLeft(Int32 step)
    {
        return ColumnShift(step * -1);
    }

    #endregion XlsxColumn Left

    #region XlsxColumn Right

    IXlsxColumn IXlsxColumn.ColumnRight()
    {
        return ColumnRight();
    }

    IXlsxColumn IXlsxColumn.ColumnRight(Int32 step)
    {
        return ColumnRight(step);
    }

    public XlsxColumn ColumnRight()
    {
        return ColumnRight(1);
    }

    public XlsxColumn ColumnRight(Int32 step)
    {
        return ColumnShift(step);
    }

    #endregion XlsxColumn Right

    public override Boolean IsEmpty()
    {
        return IsEmpty(XlsxCellsUsedOptions.AllContents);
    }

    public override Boolean IsEmpty(XlsxCellsUsedOptions options)
    {
        if (options.HasFlag(XlsxCellsUsedOptions.NormalFormats) && FormatValue is not null)
            return false;

        return base.IsEmpty(options);
    }

    public override Boolean IsEntireRow()
    {
        return false;
    }

    public override Boolean IsEntireColumn()
    {
        return true;
    }
}
