#nullable disable

using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges;

/// <summary>
/// The direct constructor should only be used in <see cref="XlsxWorksheet.RangeFactory"/>.
/// </summary>
internal class XlsxRangeColumn(XlsxRangeAddress rangeAddress) : XlsxRangeBase(rangeAddress), IXlsxRangeColumn
{
    internal override XlsxCellFormat Format => XlsxCellFormat.ForRange(Worksheet, RangeAddress);

    #region IXlsxRangeColumn Members

    IXlsxCell IXlsxRangeColumn.Cell(int rowNumber)
    {
        return Cell(rowNumber);
    }

    IXlsxCells IXlsxRangeColumn.Cells(string cellsInColumn) => Cells(cellsInColumn);

    public override XlsxCells Cells(string cellsInColumn)
    {
        var retVal = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.AllContents);
        var rangePairs = cellsInColumn.Split(',');
        foreach (string pair in rangePairs)
            retVal.Add(Range(pair.Trim()).RangeAddress);
        return retVal;
    }

    public IXlsxCells Cells(int firstRow, int lastRow)
    {
        return Cells(firstRow + ":" + lastRow);
    }

    public void Delete()
    {
        Delete(true);
    }

    internal void Delete(bool deleteTableField)
    {
        if (deleteTableField && IsTableColumn())
        {
            var table = Table as XlsxTable;
            if (!Cell(1).Value.TryGetText(out var firstCellValue))
                throw new InvalidOperationException("Top cell doesn't contain a text.");

            if (!table.FieldNames.ContainsKey(firstCellValue))
                throw new InvalidOperationException($"Field {firstCellValue} not found.");

            XlsxTableField field = table.Fields.Cast<XlsxTableField>().Single(f => f.Name == firstCellValue);
            field.Delete(false);
        }

        Delete(XlsxShiftDeletedCells.ShiftCellsLeft);
    }

    public IXlsxCells InsertCellsAbove(int numberOfRows)
    {
        return InsertCellsAbove(numberOfRows, false);
    }

    public IXlsxCells InsertCellsAbove(int numberOfRows, bool expandRange)
    {
        return InsertRowsAbove(numberOfRows, expandRange).Cells();
    }

    public IXlsxCells InsertCellsBelow(int numberOfRows)
    {
        return InsertCellsBelow(numberOfRows, true);
    }

    public IXlsxCells InsertCellsBelow(int numberOfRows, bool expandRange)
    {
        return InsertRowsBelow(numberOfRows, expandRange).Cells();
    }

    public int CellCount()
    {
        return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
    }

    public IXlsxRangeColumn Sort(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
    {
        base.Sort(1, sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    public IXlsxRangeColumn CopyTo(IXlsxCell target)
    {
        base.CopyTo((XlsxCell)target);

        int lastRowNumber = target.Address.RowNumber + RowCount() - 1;
        if (lastRowNumber > XlsxHelper.MaxRowNumber)
            lastRowNumber = XlsxHelper.MaxRowNumber;
        int lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
        if (lastColumnNumber > XlsxHelper.MaxColumnNumber)
            lastColumnNumber = XlsxHelper.MaxColumnNumber;

        return target.Worksheet.Range(
            target.Address.RowNumber,
            target.Address.ColumnNumber,
            lastRowNumber,
            lastColumnNumber)
            .Column(1);
    }

    public new IXlsxRangeColumn CopyTo(IXlsxRangeBase target)
    {
        base.CopyTo(target);

        int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
        if (lastRowNumber > XlsxHelper.MaxRowNumber)
            lastRowNumber = XlsxHelper.MaxRowNumber;
        int lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
        if (lastColumnNumber > XlsxHelper.MaxColumnNumber)
            lastColumnNumber = XlsxHelper.MaxColumnNumber;

        return target.Worksheet.Range(
            target.RangeAddress.FirstAddress.RowNumber,
            target.RangeAddress.FirstAddress.ColumnNumber,
            lastRowNumber,
            lastColumnNumber)
            .Column(1);
    }

    public IXlsxRangeColumn Column(int start, int end)
    {
        return Range(start, end).FirstColumn();
    }

    public IXlsxRangeColumn Column(IXlsxCell start, IXlsxCell end)
    {
        return Column(start.Address.RowNumber, end.Address.RowNumber);
    }

    public IXlsxRangeColumns Columns(string columns)
    {
        var retVal = new XlsxRangeColumns(Worksheet);
        var rowPairs = columns.Split(',');
        foreach (string trimmedPair in rowPairs.Select(pair => pair.Trim()))
        {
            string firstRow;
            string lastRow;
            if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
            {
                var rowRange = trimmedPair.Split(':', '-');

                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = trimmedPair;
                lastRow = trimmedPair;
            }

            retVal.Add(Range(firstRow, lastRow).FirstColumn());
        }

        return retVal;
    }

    public IXlsxColumn WorksheetColumn()
    {
        return Worksheet.Column(RangeAddress.FirstAddress.ColumnNumber);
    }

    #endregion IXlsxRangeColumn Members

    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.RangeColumn; }
    }

    public XlsxCell Cell(int row)
    {
        return Cell(row, 1);
    }

    internal override void WorksheetRangeShiftedColumns(XlsxRange range, int columnsShifted)
    {
        RangeAddress = (XlsxRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
    }

    internal override void WorksheetRangeShiftedRows(XlsxRange range, int rowsShifted)
    {
        RangeAddress = (XlsxRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
    }

    public XlsxRange Range(int firstRow, int lastRow)
    {
        return Range(firstRow, 1, lastRow, 1);
    }

    public override XlsxRange Range(string rangeAddressStr)
    {
        string rangeAddressToUse;
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

    public int CompareTo(XlsxRangeColumn otherColumn, IXlsxSortElements rowsToSort)
    {
        foreach (IXlsxSortElement e in rowsToSort)
        {
            XlsxCell thisCell = Cell(e.ElementNumber);
            XlsxCell otherCell = otherColumn.Cell(e.ElementNumber);
            int comparison;
            bool thisCellIsBlank = thisCell.IsEmpty();
            bool otherCellIsBlank = otherCell.IsEmpty();
            if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
            {
                if (thisCellIsBlank && otherCellIsBlank)
                    comparison = 0;
                else
                {
                    if (thisCellIsBlank)
                        comparison = e.SortOrder == XlsxSortOrder.Ascending ? 1 : -1;
                    else
                        comparison = e.SortOrder == XlsxSortOrder.Ascending ? -1 : 1;
                }
            }
            else
            {
                if (thisCell.DataType == otherCell.DataType)
                {
                    if (thisCell.DataType == XlsxDataType.Blank)
                        comparison = 0;
                    else if (thisCell.DataType == XlsxDataType.Boolean)
                        comparison = thisCell.GetBoolean().CompareTo(otherCell.GetBoolean());
                    else if (thisCell.DataType == XlsxDataType.Text)
                    {
                        comparison = e.MatchCase
                                         ? thisCell.GetText().CompareTo(otherCell.GetText())
                                         : string.Compare(thisCell.GetText(), otherCell.GetText(), true);
                    }
                    else if (thisCell.DataType == XlsxDataType.Error)
                        comparison = 0; // Errors are incomparable
                    else
                        comparison = thisCell.CachedValue.GetUnifiedNumber().CompareTo(thisCell.CachedValue.GetUnifiedNumber());
                }
                else if (e.MatchCase)
                    comparison = string.Compare(thisCell.GetString(), otherCell.GetString(), true);
                else
                    comparison = thisCell.GetString().CompareTo(otherCell.GetString());
            }

            if (comparison != 0)
                return e.SortOrder == XlsxSortOrder.Ascending ? comparison : comparison * -1;
        }

        return 0;
    }

    private XlsxRangeColumn ColumnShift(int columnsToShift)
    {
        int columnNumber = ColumnNumber() + columnsToShift;
        return Worksheet.Range(
            RangeAddress.FirstAddress.RowNumber,
            columnNumber,
            RangeAddress.LastAddress.RowNumber,
            columnNumber).FirstColumn();
    }

    #region XlsxRangeColumn Left

    IXlsxRangeColumn IXlsxRangeColumn.ColumnLeft()
    {
        return ColumnLeft();
    }

    IXlsxRangeColumn IXlsxRangeColumn.ColumnLeft(int step)
    {
        return ColumnLeft(step);
    }

    public XlsxRangeColumn ColumnLeft()
    {
        return ColumnLeft(1);
    }

    public XlsxRangeColumn ColumnLeft(int step)
    {
        return ColumnShift(step * -1);
    }

    #endregion XlsxRangeColumn Left

    #region XlsxRangeColumn Right

    IXlsxRangeColumn IXlsxRangeColumn.ColumnRight()
    {
        return ColumnRight();
    }

    IXlsxRangeColumn IXlsxRangeColumn.ColumnRight(int step)
    {
        return ColumnRight(step);
    }

    public XlsxRangeColumn ColumnRight()
    {
        return ColumnRight(1);
    }

    public XlsxRangeColumn ColumnRight(int step)
    {
        return ColumnShift(step);
    }

    #endregion XlsxRangeColumn Right

    public IXlsxTable AsTable()
    {
        if (IsTableColumn())
            throw new InvalidOperationException("This column is already part of a table.");

        return AsRange().AsTable();
    }

    public IXlsxTable AsTable(string name)
    {
        if (IsTableColumn())
            throw new InvalidOperationException("This column is already part of a table.");

        return AsRange().AsTable(name);
    }

    public IXlsxTable CreateTable()
    {
        if (IsTableColumn())
            throw new InvalidOperationException("This column is already part of a table.");

        return AsRange().CreateTable();
    }

    public IXlsxTable CreateTable(string name)
    {
        if (IsTableColumn())
            throw new InvalidOperationException("This column is already part of a table.");

        return AsRange().CreateTable(name);
    }

    public new IXlsxRangeColumn Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }


    public IXlsxRangeColumn ColumnUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents)
    {
        return Column((this as IXlsxRangeBase).FirstCellUsed(options),
                      (this as IXlsxRangeBase).LastCellUsed(options));
    }

    internal IXlsxTable Table { get; set; }

    public bool IsTableColumn()
    {
        return Table != null;
    }
}
