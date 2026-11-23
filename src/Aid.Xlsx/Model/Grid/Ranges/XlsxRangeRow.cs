using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges;

using System;
using System.Linq;

/// <summary>
/// The direct constructor should only be used in <see cref="XlsxWorksheet.RangeFactory"/>.
/// </summary>
internal class XlsxRangeRow(XlsxRangeAddress rangeAddress) : XlsxRangeBase(rangeAddress), IXlsxRangeRow
{
    internal override XlsxCellFormat Format => XlsxCellFormat.ForRange(Worksheet, RangeAddress);

    #region IXlsxRangeRow Members

    IXlsxCells IXlsxRangeRow.Cells(string cellsInRow) => Cells(cellsInRow);

    public IXlsxCell Cell(int column)
    {
        return Cell(1, column);
    }

    public override XlsxCell Cell(string columnLetter)
    {
        return Cell(1, columnLetter);
    }

    IXlsxCell IXlsxRangeRow.Cell(string columnLetter)
    {
        return Cell(columnLetter);
    }

    public void Delete()
    {
        Delete(XlsxShiftDeletedCells.ShiftCellsUp);
    }

    public IXlsxCells InsertCellsAfter(int numberOfColumns)
    {
        return InsertCellsAfter(numberOfColumns, true);
    }

    public IXlsxCells InsertCellsAfter(int numberOfColumns, bool expandRange)
    {
        return InsertColumnsAfter(numberOfColumns, expandRange).Cells();
    }

    public IXlsxCells InsertCellsBefore(int numberOfColumns)
    {
        return InsertCellsBefore(numberOfColumns, false);
    }

    public IXlsxCells InsertCellsBefore(int numberOfColumns, bool expandRange)
    {
        return InsertColumnsBefore(numberOfColumns, expandRange).Cells();
    }

    public override XlsxCells Cells(string cellsInRow)
    {
        var retVal = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.AllContents);
        var rangePairs = cellsInRow.Split(',');
        foreach (string pair in rangePairs)
            retVal.Add(Range(pair.Trim()).RangeAddress);
        return retVal;
    }

    public IXlsxCells Cells(int firstColumn, int lastColumn)
    {
        return Cells(firstColumn + ":" + lastColumn);
    }

    public IXlsxCells Cells(string firstColumn, string lastColumn)
    {
        return Cells(XlsxHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                     + XlsxHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public int CellCount()
    {
        return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
    }

    public new IXlsxRangeRow Sort()
    {
        return SortLeftToRight();
    }

    public new IXlsxRangeRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    public IXlsxRangeRow CopyTo(IXlsxCell target)
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
            .Row(1);
    }

    public new IXlsxRangeRow CopyTo(IXlsxRangeBase target)
    {
        base.CopyTo(target);
        int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
        if (lastRowNumber > XlsxHelper.MaxRowNumber)
            lastRowNumber = XlsxHelper.MaxRowNumber;
        int lastColumnNumber = target.RangeAddress.LastAddress.ColumnNumber + ColumnCount() - 1;
        if (lastColumnNumber > XlsxHelper.MaxColumnNumber)
            lastColumnNumber = XlsxHelper.MaxColumnNumber;

        return target.Worksheet.Range(
            target.RangeAddress.FirstAddress.RowNumber,
            target.RangeAddress.LastAddress.ColumnNumber,
            lastRowNumber,
            lastColumnNumber)
            .Row(1);
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
        var columnPairs = rows.Split(',');
        foreach (string trimmedPair in columnPairs.Select(pair => pair.Trim()))
        {
            string firstColumn;
            string lastColumn;
            if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
            {
                var columnRange = trimmedPair.Contains('-')
                                      ? trimmedPair.Replace('-', ':').Split(':')
                                      : trimmedPair.Split(':');
                firstColumn = columnRange[0];
                lastColumn = columnRange[1];
            }
            else
            {
                firstColumn = trimmedPair;
                lastColumn = trimmedPair;
            }

            retVal.Add(Range(firstColumn, lastColumn).FirstRow());
        }

        return retVal;
    }

    public IXlsxRow WorksheetRow()
    {
        return Worksheet.Row(RangeAddress.FirstAddress.RowNumber);
    }

    #endregion IXlsxRangeRow Members
    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.RangeRow; }
    }

    internal override void WorksheetRangeShiftedColumns(XlsxRange range, int columnsShifted)
    {
        RangeAddress = (XlsxRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
    }

    internal override void WorksheetRangeShiftedRows(XlsxRange range, int rowsShifted)
    {
        RangeAddress = (XlsxRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
    }

    public IXlsxRange Range(int firstColumn, int lastColumn)
    {
        return Range(1, firstColumn, 1, lastColumn);
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
            rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
        }
        else
            rangeAddressToUse = FixRowAddress(rangeAddressStr);

        var rangeAddress = new XlsxRangeAddress(Worksheet, rangeAddressToUse);
        return Range(rangeAddress);
    }

    public int CompareTo(XlsxRangeRow otherRow, IXlsxSortElements columnsToSort)
    {
        foreach (IXlsxSortElement e in columnsToSort)
        {
            var thisCell = (XlsxCell)Cell(e.ElementNumber);
            var otherCell = (XlsxCell)otherRow.Cell(e.ElementNumber);
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
                    comparison = thisCell.DataType switch
                    {
                        XlsxDataType.Text => e.MatchCase
                                                                     ? thisCell.GetText().CompareTo(otherCell.GetText())
                                                                     : string.Compare(thisCell.GetText(), otherCell.GetText(), true),
                        XlsxDataType.TimeSpan => thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan()),
                        XlsxDataType.DateTime => thisCell.GetDateTime().CompareTo(otherCell.GetDateTime()),
                        XlsxDataType.Number => thisCell.GetDouble().CompareTo(otherCell.GetDouble()),
                        XlsxDataType.Boolean => thisCell.GetBoolean().CompareTo(otherCell.GetBoolean()),
                        _ => throw new NotImplementedException(),
                    };
                }
                else if (thisCell.Value.IsUnifiedNumber && otherCell.Value.IsUnifiedNumber)
                    comparison = thisCell.Value.GetUnifiedNumber().CompareTo(otherCell.Value.GetUnifiedNumber());
                else if (e.MatchCase)
                    comparison = string.Compare(thisCell.GetString(), otherCell.GetString(), true);
                else
                    comparison = thisCell.GetString().CompareTo(otherCell.GetString());
            }

            if (comparison != 0)
                return e.SortOrder == XlsxSortOrder.Ascending ? comparison : -comparison;
        }

        return 0;
    }

    private XlsxRangeRow RowShift(int rowsToShift)
    {
        int rowNum = RowNumber() + rowsToShift;

        XlsxRange range = Worksheet.Range(
            rowNum,
            RangeAddress.FirstAddress.ColumnNumber,
            rowNum,
            RangeAddress.LastAddress.ColumnNumber);

        return range.FirstRow();
    }

    #region XlsxRangeRow Above

    IXlsxRangeRow IXlsxRangeRow.RowAbove()
    {
        return RowAbove();
    }

    IXlsxRangeRow IXlsxRangeRow.RowAbove(int step)
    {
        return RowAbove(step);
    }

    public XlsxRangeRow RowAbove()
    {
        return RowAbove(1);
    }

    public XlsxRangeRow RowAbove(int step)
    {
        return RowShift(step * -1);
    }

    #endregion XlsxRangeRow Above

    #region XlsxRangeRow Below

    IXlsxRangeRow IXlsxRangeRow.RowBelow()
    {
        return RowBelow();
    }

    IXlsxRangeRow IXlsxRangeRow.RowBelow(int step)
    {
        return RowBelow(step);
    }

    public XlsxRangeRow RowBelow()
    {
        return RowBelow(1);
    }

    public XlsxRangeRow RowBelow(int step)
    {
        return RowShift(step);
    }

    #endregion XlsxRangeRow Below

    public new IXlsxRangeRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }


    public IXlsxRangeRow RowUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents)
    {
        return Row((this as IXlsxRangeBase).FirstCellUsed(options),
                   (this as IXlsxRangeBase).LastCellUsed(options));
    }
}
