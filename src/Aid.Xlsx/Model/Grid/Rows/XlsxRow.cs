using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Rendering;

namespace Aid.Xlsx.Model.Grid.Rows;

internal sealed class XlsxRow : XlsxRangeBase, IXlsxRow, IXlsxFormatContainer
{
    /// <summary>
    /// Don't use directly, use properties.
    /// </summary>
    private XlRowFlags _flags;
    private double _height;
    private int _outlineLevel;

    /// <summary>
    /// The direct constructor should only be used in <see cref="XlsxWorksheet.RangeFactory"/>.
    /// </summary>
    public XlsxRow(XlsxWorksheet worksheet, int row)
        : base(XlsxRangeAddress.EntireRow(worksheet, row))
    {
        SetRowNumber(row);

        _height = worksheet.RowHeight;
    }

    internal XlsxRowArea Area => new(Worksheet.Name, RowNumber());

    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.Row; }
    }

    public bool Collapsed
    {
        get => _flags.HasFlag(XlRowFlags.Collapsed);
        set
        {
            if (value)
                _flags |= XlRowFlags.Collapsed;
            else
                _flags &= ~XlRowFlags.Collapsed;
        }
    }

    /// <summary>
    /// Distance in pixels from the bottom of the cells in the current row to the typographical
    /// baseline of the cell content if, hypothetically, the zoom level for the sheet containing
    /// this row is 100 percent and the cell has bottom-alignment formatting.
    /// </summary>
    /// <remarks>
    /// If the attribute is set, it sets customHeight to true even if the customHeight is explicitly
    /// set to false. Custom height means no auto-sizing by Excel on load, so if row has this
    /// attribute, it stops Excel from auto-sizing the height of a row to fit the content on load.
    /// </remarks>
    public double? DyDescent { get; set; }

    /// <summary>
    /// Should cells in the row display phonetic? This doesn't actually affect whether the phonetic are
    /// shown in the row, that depends entirely on the <see cref="IXlsxCell.ShowPhonetic"/> property
    /// of a cell. This property determines whether a new cell in the row will have it's phonetic turned on
    /// (and also the state of the "Show or hide phonetic" in Excel when whole row is selected).
    /// Default is <c>false</c>.
    /// </summary>
    public bool ShowPhonetic
    {
        get => _flags.HasFlag(XlRowFlags.ShowPhonetic);
        set
        {
            if (value)
                _flags |= XlRowFlags.ShowPhonetic;
            else
                _flags &= ~XlRowFlags.ShowPhonetic;
        }
    }

    public bool Loading
    {
        get => _flags.HasFlag(XlRowFlags.Loading);
        set
        {
            if (value)
                _flags |= XlRowFlags.Loading;
            else
                _flags &= ~XlRowFlags.Loading;
        }
    }

    /// <summary>
    /// Does row have an individual height or is it derived from the worksheet <see cref="XlsxWorksheet.RowHeight"/>?
    /// </summary>
    public bool HeightChanged
    {
        get => _flags.HasFlag(XlRowFlags.HeightChanged);
        private set
        {
            if (value)
                _flags |= XlRowFlags.HeightChanged;
            else
                _flags &= ~XlRowFlags.HeightChanged;
        }
    }

    #region IXlsxRow Members

    public double Height
    {
        get { return _height; }
        set
        {
            if (!Loading)
                HeightChanged = true;

            _height = value;
        }
    }

    IXlsxCells IXlsxRow.Cells(string cellsInRow) => Cells(cellsInRow);

    IXlsxCells IXlsxRow.Cells(int firstColumn, int lastColumn) => Cells(firstColumn, lastColumn);

    public void ClearHeight()
    {
        Height = Worksheet.RowHeight;
        HeightChanged = false;
    }

    public void Delete()
    {
        var rowNumber = RowNumber();
        AsRange().Delete(XlsxShiftDeletedCells.ShiftCellsUp);
        Worksheet.DeleteRow(rowNumber);
    }

    public new IXlsxRows InsertRowsBelow(int numberOfRows)
    {
        var rowNum = RowNumber();
        Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
        XlsxRange asRange = Worksheet.Row(rowNum).AsRange();
        asRange.InsertRowsBelowVoid(true, numberOfRows);

        IXlsxRows newRows = Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);
        foreach (IXlsxRow? newRow in newRows)
        {
            XlsxRow internalRow = Worksheet.Internals.RowsCollection[newRow.RowNumber()];
            internalRow._height = Height;
            internalRow.FormatValue = FormatValue; // Is within a worbook
            internalRow.Collapsed = Collapsed;
            internalRow.IsHidden = IsHidden;
            internalRow._outlineLevel = OutlineLevel;
        }

        return newRows;
    }

    public new IXlsxRows InsertRowsAbove(int numberOfRows)
    {
        var rowNum = RowNumber();
        if (rowNum > 1)
        {
            return Worksheet.Row(rowNum - 1).InsertRowsBelow(numberOfRows);
        }

        Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
        XlsxRange asRange = Worksheet.Row(rowNum).AsRange();
        asRange.InsertRowsAboveVoid(true, numberOfRows);

        return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
    }

    public new IXlsxRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXlsxCell Cell(int columnNumber)
    {
        return Cell(1, columnNumber);
    }

    public override XlsxCell Cell(string columnLetter)
    {
        return Cell(1, columnLetter);
    }

    IXlsxCell IXlsxRow.Cell(string columnLetter)
    {
        return Cell(columnLetter);
    }

    public override IXlsxCells Cells()
    {
        return Cells(true, XlsxCellsUsedOptions.All);
    }

    public override XlsxCells Cells(bool usedCellsOnly)
    {
        if (usedCellsOnly)
            return Cells(true, XlsxCellsUsedOptions.AllContents);
        else
            return Cells(FirstCellUsed().Address.ColumnNumber, LastCellUsed().Address.ColumnNumber);
    }

    public override XlsxCells Cells(string cellsInRow)
    {
        var retVal = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.AllContents);
        var rangePairs = cellsInRow.Split(',');
        foreach (var pair in rangePairs)
            retVal.Add(Range(pair.Trim()).RangeAddress);
        return retVal;
    }

    public XlsxCells Cells(int firstColumn, int lastColumn)
    {
        return Cells(firstColumn + ":" + lastColumn);
    }

    public IXlsxCells Cells(string firstColumn, string lastColumn)
    {
        return Cells(XlsxHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                     + XlsxHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public IXlsxRow AdjustToContents(int startColumn)
    {
        return AdjustToContents(startColumn, XlsxHelper.MaxColumnNumber);
    }

    public IXlsxRow AdjustToContents(int startColumn, int endColumn)
    {
        return AdjustToContents(startColumn, endColumn, 0, double.MaxValue);
    }

    public IXlsxRow AdjustToContents(double minHeight, double maxHeight)
    {
        return AdjustToContents(1, XlsxHelper.MaxColumnNumber, minHeight, maxHeight);
    }

    public IXlsxRow AdjustToContents(int startColumn, double minHeight, double maxHeight)
    {
        return AdjustToContents(startColumn, XlsxHelper.MaxColumnNumber, minHeight, maxHeight);
    }

    public IXlsxRow AdjustToContents(int startColumn, int endColumn, double minHeightPt, double maxHeightPt)
    {
        IXlsxGraphicEngine engine = Worksheet.Workbook.GraphicEngine;
        var dpi = new Dpi(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);

        var rowHeightPx = CalculateMinRowHeight(startColumn, endColumn, engine, dpi);

        var rowHeightPt = XlsxHelper.PixelsToPoints(rowHeightPx, dpi.Y);
        if (rowHeightPt <= 0)
            rowHeightPt = Worksheet.RowHeight;

        if (minHeightPt > rowHeightPt)
            rowHeightPt = minHeightPt;

        if (maxHeightPt < rowHeightPt)
            rowHeightPt = maxHeightPt;

        Height = rowHeightPt;

        return this;
    }

    private int CalculateMinRowHeight(int startColumn, int endColumn, IXlsxGraphicEngine engine, Dpi dpi)
    {
        var glyphs = new List<GlyphBox>();
        var rowHeightPx = 0;
        foreach (XlsxCell cell in Row(startColumn, endColumn).CellsUsed().Cast<XlsxCell>())
        {
            // Clear maintains capacity -> reduce need for GC
            glyphs.Clear();

            if (cell.IsMerged())
                continue;

            XlsxCellFormatValue cellFormat = Worksheet.GetStyleValue(cell.Point);

            cell.GetGlyphBoxes(engine, dpi, glyphs);
            var cellHeightPx = (int)Math.Ceiling(GetContentHeight(cellFormat.Alignment.TextRotation.Value, glyphs));

            rowHeightPx = Math.Max(cellHeightPx, rowHeightPx);
        }

        return rowHeightPx;
    }

    private static double GetContentHeight(int textRotationDeg, List<GlyphBox> glyphs)
    {
        if (textRotationDeg == 0)
        {
            var textHeight = 0d;
            var lineMaxHeight = 0d;
            foreach (GlyphBox glyph in glyphs)
            {
                if (!glyph.IsLineBreak)
                {
                    var cellHeightPx = glyph.LineHeight;
                    lineMaxHeight = Math.Max(cellHeightPx, lineMaxHeight);
                }
                else
                {
                    // At the end of each line, add height of the line to total height.
                    textHeight += lineMaxHeight;
                    lineMaxHeight = 0d;
                }
            }

            // If the last line ends without EOL, it must be also counted
            textHeight += lineMaxHeight;

            return textHeight;
        }
        else if (textRotationDeg == 255)
        {
            // Glyphs are vertically aligned.
            var textHeight = glyphs.Sum(static g => g.LineHeight);
            return textHeight;
        }
        else
        {
            // Rotated text
            var width = 0d;
            var height = 0d;
            foreach (GlyphBox glyph in glyphs)
            {
                width += glyph.AdvanceWidth;
                height = Math.Max(glyph.LineHeight, height);
            }

            var projectedWidth = Math.Sin(XlsxHelper.DegToRad(textRotationDeg)) * width;
            var projectedHeight = Math.Cos(XlsxHelper.DegToRad(textRotationDeg)) * height;
            return projectedWidth + projectedHeight;
        }
    }

    public IXlsxRow Hide()
    {
        IsHidden = true;
        return this;
    }

    public IXlsxRow Unhide()
    {
        IsHidden = false;
        return this;
    }

    public bool IsHidden
    {
        get => _flags.HasFlag(XlRowFlags.IsHidden);
        set
        {
            if (value)
                _flags |= XlRowFlags.IsHidden;
            else
                _flags &= ~XlRowFlags.IsHidden;
        }
    }

    public int OutlineLevel
    {
        get { return _outlineLevel; }
        set
        {
            if (value < 0 || value > 8)
                throw new ArgumentOutOfRangeException(nameof(value), "Outline level must be between 0 and 8.");

            Worksheet.IncrementColumnOutline(value);
            Worksheet.DecrementColumnOutline(_outlineLevel);
            _outlineLevel = value;
        }
    }

    public IXlsxRow Group()
    {
        return Group(false);
    }

    public IXlsxRow Group(int outlineLevel)
    {
        return Group(outlineLevel, false);
    }

    public IXlsxRow Ungroup()
    {
        return Ungroup(false);
    }

    public IXlsxRow Group(bool collapse)
    {
        if (OutlineLevel < 8)
            OutlineLevel += 1;

        Collapsed = collapse;
        return this;
    }

    public IXlsxRow Group(int outlineLevel, bool collapse)
    {
        OutlineLevel = outlineLevel;
        Collapsed = collapse;
        return this;
    }

    public IXlsxRow Ungroup(bool ungroupFromAll)
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

    public IXlsxRow Collapse()
    {
        Collapsed = true;
        return Hide();
    }

    public IXlsxRow Expand()
    {
        Collapsed = false;
        return Unhide();
    }

    public int CellCount()
    {
        return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
    }

    public new IXlsxRow Sort()
    {
        return SortLeftToRight();
    }

    public new IXlsxRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false,
                                      bool ignoreBlanks = true)
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    IXlsxRangeRow IXlsxRow.CopyTo(IXlsxCell target)
    {
        IXlsxRange copy = AsRange().CopyTo(target);
        return copy.Row(1);
    }

    IXlsxRangeRow IXlsxRow.CopyTo(IXlsxRangeBase target)
    {
        IXlsxRange copy = AsRange().CopyTo(target);
        return copy.Row(1);
    }

    public IXlsxRow CopyTo(IXlsxRow row)
    {
        row.Clear();
        var newRow = (XlsxRow)row;
        newRow._height = _height;
        newRow.HeightChanged = HeightChanged;
        newRow.IsHidden = IsHidden;
        if (FormatValue is not null)
        {
            newRow.FormatValue = newRow.Worksheet.Workbook.Styles.GetRegisteredCellFormat(GetFormat());
        }

        AsRange().CopyTo(row);

        return newRow;
    }

    public IXlsxRangeRow Row(int start, int end)
    {
        return Range(1, start, 1, end).Row(1);
    }

    public IXlsxRangeRow Row(IXlsxCell start, IXlsxCell end)
    {
        return Row(start.Address.ColumnNumber, end.Address.ColumnNumber);
    }

    public IXlsxRangeRows Rows(string rows)
    {
        var retVal = new XlsxRangeRows(Worksheet);
        var rowPairs = rows.Split(',');
        foreach (var pair in rowPairs)
            AsRange().Rows(pair.Trim()).ForEach(retVal.Add);

        return retVal;
    }

    public IXlsxRow AddHorizontalPageBreak()
    {
        Worksheet.PageSetup.AddHorizontalPageBreak(RowNumber());
        return this;
    }

    public IXlsxRangeRow RowUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents)
    {
        return Row((this as IXlsxRangeBase).FirstCellUsed(options),
            (this as IXlsxRangeBase).LastCellUsed(options));
    }

    #endregion IXlsxRow Members


    #region IXlsxFormatContainer

    /// <remarks>
    /// Format of a row or <c>null</c> for not defined format.
    /// </remarks>
    /// <inheritdoc cref="IXlsxFormatContainer.FormatValue"/>
    public XlsxCellFormatValue? FormatValue { get; set; }

    internal override XlsxCellFormat Format => XlsxCellFormat.ForRow(this);

    private XlsxCellFormatValue GetFormat()
    {
        return FormatValue ?? Worksheet.Workbook.Styles.DefaultFormat;
    }

    #endregion

    public override XlsxRange AsRange()
    {
        return Range(1, 1, 1, XlsxHelper.MaxColumnNumber);
    }

    internal override void WorksheetRangeShiftedColumns(XlsxRange range, int columnsShifted)
    {
        //do nothing
    }

    internal override void WorksheetRangeShiftedRows(XlsxRange range, int rowsShifted)
    {
        // rows are shifted by XlsxRowCollection
    }

    internal void SetRowNumber(int row)
    {
        RangeAddress = new XlsxRangeAddress(
            new XlsxAddress(Worksheet, row, 1, RangeAddress.FirstAddress.FixedRow,
                          RangeAddress.FirstAddress.FixedColumn),
            new XlsxAddress(Worksheet,
                          row,
                          XlsxHelper.MaxColumnNumber,
                          RangeAddress.LastAddress.FixedRow,
                          RangeAddress.LastAddress.FixedColumn));
    }

    public override XlsxRange Range(string rangeAddressStr)
    {
        string rangeAddressToUse;
        if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
        {
            if (rangeAddressStr.Contains('-'))
                rangeAddressStr = rangeAddressStr.Replace('-', ':');

            var arrRange = rangeAddressStr.Split(':');
            var firstPart = arrRange[0];
            var secondPart = arrRange[1];
            rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
        }
        else
            rangeAddressToUse = FixRowAddress(rangeAddressStr);

        var rangeAddress = new XlsxRangeAddress(Worksheet, rangeAddressToUse);
        return Range(rangeAddress);
    }

    public IXlsxRow AdjustToContents()
    {
        return AdjustToContents(1);
    }

    private XlsxRow RowShift(int rowsToShift)
    {
        return Worksheet.Row(RowNumber() + rowsToShift);
    }

    #region XlsxRow Above

    IXlsxRow IXlsxRow.RowAbove()
    {
        return RowAbove();
    }

    IXlsxRow IXlsxRow.RowAbove(int step)
    {
        return RowAbove(step);
    }

    public XlsxRow RowAbove()
    {
        return RowAbove(1);
    }

    public XlsxRow RowAbove(int step)
    {
        return RowShift(step * -1);
    }

    #endregion XlsxRow Above

    #region XlsxRow Below

    IXlsxRow IXlsxRow.RowBelow()
    {
        return RowBelow();
    }

    IXlsxRow IXlsxRow.RowBelow(int step)
    {
        return RowBelow(step);
    }

    public XlsxRow RowBelow()
    {
        return RowBelow(1);
    }

    public XlsxRow RowBelow(int step)
    {
        return RowShift(step);
    }

    #endregion XlsxRow Below

    public override bool IsEmpty()
    {
        return IsEmpty(XlsxCellsUsedOptions.AllContents);
    }

    public override bool IsEmpty(XlsxCellsUsedOptions options)
    {
        if (options.HasFlag(XlsxCellsUsedOptions.NormalFormats) && FormatValue is not null)
            return false;

        return base.IsEmpty(options);
    }

    public override bool IsEntireRow()
    {
        return true;
    }

    public override bool IsEntireColumn()
    {
        return false;
    }

    /// <summary>
    /// Flag enum to save space, instead of wasting byte for each flag.
    /// </summary>
    [Flags]
    private enum XlRowFlags : byte
    {
        Collapsed = 1,
        IsHidden = 2,
        ShowPhonetic = 4,
        HeightChanged = 8,
        Loading = 16
    }
}
