#nullable disable

using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Tables;

namespace Aid.Xlsx.Model.Grid.Ranges;

internal class XlsxRange(XlsxRangeAddress rangeAddress) : XlsxRangeBase(rangeAddress), IXlsxRange
{
    internal override XlsxCellFormat Format => XlsxCellFormat.ForRange(Worksheet, RangeAddress);

    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.Range; }
    }

    #region IXlsxRange Members

    IXlsxRangeRow IXlsxRange.Row(int row)
    {
        return Row(row);
    }

    IXlsxRangeColumn IXlsxRange.Column(int columnNumber)
    {
        return Column(columnNumber);
    }

    IXlsxRangeColumn IXlsxRange.Column(string columnLetter)
    {
        return Column(columnLetter);
    }

    public virtual IXlsxRangeColumns Columns(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        var retVal = new XlsxRangeColumns(Worksheet);
        int columnCount = ColumnCount();
        for (int c = 1; c <= columnCount; c++)
        {
            XlsxRangeColumn column = Column(c);
            if (predicate == null || predicate(column))
                retVal.Add(column);
        }
        return retVal;
    }

    public virtual IXlsxRangeColumns Columns(int firstColumn, int lastColumn)
    {
        var retVal = new XlsxRangeColumns(Worksheet);

        for (int co = firstColumn; co <= lastColumn; co++)
            retVal.Add(Column(co));
        return retVal;
    }

    public virtual IXlsxRangeColumns Columns(string firstColumn, string lastColumn)
    {
        return Columns(XlsxHelper.GetColumnNumberFromLetter(firstColumn),
                       XlsxHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public virtual IXlsxRangeColumns Columns(string columns)
    {
        var retVal = new XlsxRangeColumns(Worksheet);
        var columnPairs = columns.Split(',');
        foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
        {
            string firstColumn;
            string lastColumn;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                string[] columnRange = XlsxHelper.SplitRange(tPair);

                firstColumn = columnRange[0];
                lastColumn = columnRange[1];
            }
            else
            {
                firstColumn = tPair;
                lastColumn = tPair;
            }

            if (int.TryParse(firstColumn, out int tmp))
            {
                foreach (IXlsxRangeColumn col in Columns(int.Parse(firstColumn), int.Parse(lastColumn)))
                    retVal.Add(col);
            }
            else
            {
                foreach (IXlsxRangeColumn col in Columns(firstColumn, lastColumn))
                    retVal.Add(col);
            }
        }
        return retVal;
    }

    IXlsxCell IXlsxRange.Cell(int row, int column)
    {
        return Cell(row, column);
    }

    IXlsxCell IXlsxRange.Cell(string cellAddressInRange)
    {
        return Cell(cellAddressInRange);
    }

    IXlsxCell IXlsxRange.Cell(int row, string column)
    {
        return Cell(row, column);
    }

    IXlsxCell IXlsxRange.Cell(IXlsxAddress cellAddressInRange)
    {
        return Cell(cellAddressInRange);
    }

    IXlsxRange IXlsxRange.Range(IXlsxRangeAddress rangeAddress)
    {
        return Range(rangeAddress);
    }

    IXlsxRange IXlsxRange.Range(string rangeAddress)
    {
        return Range(rangeAddress);
    }

    IXlsxRange IXlsxRange.Range(IXlsxCell firstCell, IXlsxCell lastCell)
    {
        return Range(firstCell, lastCell);
    }

    IXlsxRange IXlsxRange.Range(string firstCellAddress, string lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXlsxRange IXlsxRange.Range(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXlsxRange IXlsxRange.Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        return Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
    }

    IXlsxRanges IXlsxRange.Ranges(string ranges) => Ranges(ranges);

    public IXlsxRangeRows Rows(Func<IXlsxRangeRow, bool> predicate = null)
    {
        var retVal = new XlsxRangeRows(Worksheet);
        int rowCount = RowCount();
        for (int r = 1; r <= rowCount; r++)
        {
            XlsxRangeRow row = Row(r);
            if (predicate == null || predicate(row))
                retVal.Add(Row(r));
        }
        return retVal;
    }

    public IXlsxRangeRows Rows(int firstRow, int lastRow)
    {
        var retVal = new XlsxRangeRows(Worksheet);

        for (int ro = firstRow; ro <= lastRow; ro++)
            retVal.Add(Row(ro));
        return retVal;
    }

    public IXlsxRangeRows Rows(string rows)
    {
        var retVal = new XlsxRangeRows(Worksheet);
        var rowPairs = rows.Split(',');
        foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
        {
            string firstRow;
            string lastRow;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                string[] rowRange = XlsxHelper.SplitRange(tPair);

                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = tPair;
                lastRow = tPair;
            }
            foreach (IXlsxRangeRow row in Rows(int.Parse(firstRow), int.Parse(lastRow)))
                retVal.Add(row);
        }
        return retVal;
    }

    public void Transpose(XlsxTransposeOptions transposeOption)
    {
        int rowCount = RowCount();
        int columnCount = ColumnCount();
        int squareSide = rowCount > columnCount ? rowCount : columnCount;

        XlsxCell firstCell = FirstCell();

        MoveOrClearForTranspose(transposeOption, rowCount, columnCount);
        TransposeMerged(squareSide);
        TransposeRange(squareSide);
        RangeAddress = new XlsxRangeAddress(
            RangeAddress.FirstAddress,
            new XlsxAddress(Worksheet,
                          firstCell.Address.RowNumber + columnCount - 1,
                          firstCell.Address.ColumnNumber + rowCount - 1,
                          RangeAddress.LastAddress.FixedRow,
                          RangeAddress.LastAddress.FixedColumn));

        if (rowCount > columnCount)
        {
            XlsxRange rng = Worksheet.Range(
                RangeAddress.LastAddress.RowNumber + 1,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber + (rowCount - columnCount),
                RangeAddress.LastAddress.ColumnNumber);
            rng.Delete(XlsxShiftDeletedCells.ShiftCellsUp);
        }
        else if (columnCount > rowCount)
        {
            XlsxRange rng = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber + 1,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber + (columnCount - rowCount));
            rng.Delete(XlsxShiftDeletedCells.ShiftCellsLeft);
        }

        foreach (IXlsxCell c in Range(1, 1, columnCount, rowCount).Cells())
        {
            XlsxBorderFormatValue border = ((XlsxCell)c).GetFormat().Border;
            c.Style.Border.TopBorder = border.Left.Style;
            c.Style.Border.TopBorderColor = border.Left.Color;
            c.Style.Border.LeftBorder = border.Top.Style;
            c.Style.Border.LeftBorderColor = border.Top.Color;
            c.Style.Border.RightBorder = border.Bottom.Style;
            c.Style.Border.RightBorderColor = border.Bottom.Color;
            c.Style.Border.BottomBorder = border.Right.Style;
            c.Style.Border.BottomBorderColor = border.Right.Color;
        }
    }

    public IXlsxTable AsTable()
    {
        return Worksheet.Table(this, false);
    }

    public IXlsxTable AsTable(string name)
    {
        return Worksheet.Table(this, name, false);
    }

    IXlsxTable IXlsxRange.CreateTable()
    {
        return CreateTable();
    }

    public XlsxTable CreateTable()
    {
        return (XlsxTable)Worksheet.Table(this, true, true);
    }

    IXlsxTable IXlsxRange.CreateTable(string name)
    {
        return CreateTable(name);
    }

    public XlsxTable CreateTable(string name)
    {
        return (XlsxTable)Worksheet.Table(this, name, true, true);
    }

    public IXlsxTable CreateTable(string name, bool setAutofilter)
    {
        return Worksheet.Table(this, name, true, setAutofilter);
    }

    public IXlsxRange CopyTo(IXlsxCell target)
    {
        base.CopyTo((XlsxCell)target);

        int lastRowNumber = target.Address.RowNumber + RowCount() - 1;
        if (lastRowNumber > XlsxHelper.MaxRowNumber)
            lastRowNumber = XlsxHelper.MaxRowNumber;
        int lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
        if (lastColumnNumber > XlsxHelper.MaxColumnNumber)
            lastColumnNumber = XlsxHelper.MaxColumnNumber;

        return target.Worksheet.Range(target.Address.RowNumber,
                                      target.Address.ColumnNumber,
                                      lastRowNumber,
                                      lastColumnNumber);
    }

    public new IXlsxRange CopyTo(IXlsxRangeBase target)
    {
        base.CopyTo(target);

        int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
        if (lastRowNumber > XlsxHelper.MaxRowNumber)
            lastRowNumber = XlsxHelper.MaxRowNumber;
        int lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
        if (lastColumnNumber > XlsxHelper.MaxColumnNumber)
            lastColumnNumber = XlsxHelper.MaxColumnNumber;

        return target.Worksheet.Range(target.RangeAddress.FirstAddress.RowNumber,
                                      target.RangeAddress.FirstAddress.ColumnNumber,
                                      lastRowNumber,
                                      lastColumnNumber);
    }

    public new IXlsxRange Sort()
    {
        return base.Sort().AsRange();
    }

    public new IXlsxRange Sort(string columnsToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
    {
        return base.Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    public new IXlsxRange Sort(int columnToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
    {
        return base.Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    public new IXlsxRange SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
    {
        return base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    #endregion IXlsxRange Members

    internal override void WorksheetRangeShiftedColumns(XlsxRange range, int columnsShifted)
    {
        RangeAddress = (XlsxRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
    }

    internal override void WorksheetRangeShiftedRows(XlsxRange range, int rowsShifted)
    {
        RangeAddress = (XlsxRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
    }

    IXlsxRangeColumn IXlsxRange.FirstColumn(Func<IXlsxRangeColumn, bool> predicate)
    {
        return FirstColumn(predicate);
    }

    internal XlsxRangeColumn FirstColumn(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        if (predicate == null)
            return Column(1);

        int columnCount = ColumnCount();
        for (int c = 1; c <= columnCount; c++)
        {
            XlsxRangeColumn column = Column(c);
            if (predicate(column)) return column;
        }

        return null;
    }

    IXlsxRangeColumn IXlsxRange.LastColumn(Func<IXlsxRangeColumn, bool> predicate)
    {
        return LastColumn(predicate);
    }

    internal XlsxRangeColumn LastColumn(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        int columnCount = ColumnCount();
        if (predicate == null)
            return Column(columnCount);

        for (int c = columnCount; c >= 1; c--)
        {
            XlsxRangeColumn column = Column(c);
            if (predicate(column)) return column;
        }

        return null;
    }

    IXlsxRangeColumn IXlsxRange.FirstColumnUsed(Func<IXlsxRangeColumn, bool> predicate)
    {
        return FirstColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeColumn FirstColumnUsed(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        return FirstColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeColumn IXlsxRange.FirstColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, bool> predicate)
    {
        return FirstColumnUsed(options, predicate);
    }

    internal XlsxRangeColumn FirstColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, bool> predicate = null)
    {
        if (predicate == null)
        {
            int firstColumnUsed = Worksheet.Internals.CellsCollection.FirstColumnUsed(
                Area.FromRangeAddress(RangeAddress),
                options);

            return firstColumnUsed == 0 ? null : Column(firstColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        int columnCount = ColumnCount();
        for (int co = 1; co <= columnCount; co++)
        {
            XlsxRangeColumn column = Column(co);

            if (!column.IsEmpty(options) && predicate(column))
                return column;
        }
        return null;
    }

    IXlsxRangeColumn IXlsxRange.LastColumnUsed(Func<IXlsxRangeColumn, bool> predicate)
    {
        return LastColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeColumn LastColumnUsed(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        return LastColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeColumn IXlsxRange.LastColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, bool> predicate)
    {
        return LastColumnUsed(options, predicate);
    }

    internal XlsxRangeColumn LastColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, bool> predicate = null)
    {
        if (predicate == null)
        {
            int lastColumnUsed = Worksheet.Internals.CellsCollection.LastColumnUsed(
                Area.FromRangeAddress(RangeAddress),
                options);

            return lastColumnUsed == 0 ? null : Column(lastColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        int columnCount = ColumnCount();
        for (int co = columnCount; co >= 1; co--)
        {
            XlsxRangeColumn column = Column(co);

            if (!column.IsEmpty(options) && predicate(column))
                return column;
        }
        return null;
    }

    IXlsxRangeRow IXlsxRange.FirstRow(Func<IXlsxRangeRow, bool> predicate)
    {
        return FirstRow(predicate);
    }

    public XlsxRangeRow FirstRow(Func<IXlsxRangeRow, bool> predicate = null)
    {
        if (predicate == null)
            return Row(1);

        int rowCount = RowCount();
        for (int ro = 1; ro <= rowCount; ro++)
        {
            XlsxRangeRow row = Row(ro);
            if (predicate(row)) return row;
        }

        return null;
    }

    IXlsxRangeRow IXlsxRange.LastRow(Func<IXlsxRangeRow, bool> predicate)
    {
        return LastRow(predicate);
    }

    public XlsxRangeRow LastRow(Func<IXlsxRangeRow, bool> predicate = null)
    {
        int rowCount = RowCount();
        if (predicate == null)
            return Row(rowCount);

        for (int ro = rowCount; ro >= 1; ro--)
        {
            XlsxRangeRow row = Row(ro);
            if (predicate(row)) return row;
        }

        return null;
    }

    IXlsxRangeRow IXlsxRange.FirstRowUsed(Func<IXlsxRangeRow, bool> predicate)
    {
        return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeRow FirstRowUsed(Func<IXlsxRangeRow, bool> predicate = null)
    {
        return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeRow IXlsxRange.FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, bool> predicate)
    {
        return FirstRowUsed(options, predicate);
    }

    internal XlsxRangeRow FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, bool> predicate = null)
    {
        if (predicate == null)
        {
            int rowFromCells = Worksheet.Internals.CellsCollection.FirstRowUsed(
                Area.FromRangeAddress(RangeAddress), options);

            return rowFromCells == 0 ? null : Row(rowFromCells - RangeAddress.FirstAddress.RowNumber + 1);
        }

        int rowCount = RowCount();
        for (int ro = 1; ro <= rowCount; ro++)
        {
            XlsxRangeRow row = Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }
        return null;
    }

    IXlsxRangeRow IXlsxRange.LastRowUsed(Func<IXlsxRangeRow, bool> predicate)
    {
        return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeRow LastRowUsed(Func<IXlsxRangeRow, bool> predicate = null)
    {
        return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeRow IXlsxRange.LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, bool> predicate)
    {
        return LastRowUsed(options, predicate);
    }

    internal XlsxRangeRow LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, bool> predicate = null)
    {
        if (predicate == null)
        {
            int lastRowUsed = Worksheet.Internals.CellsCollection.LastRowUsed(
                Area.FromRangeAddress(RangeAddress), options);

            return lastRowUsed == 0 ? null : Row(lastRowUsed - RangeAddress.FirstAddress.RowNumber + 1);
        }

        int rowCount = RowCount();
        for (int ro = rowCount; ro >= 1; ro--)
        {
            XlsxRangeRow row = Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }
        return null;
    }

    IXlsxRangeRows IXlsxRange.RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, bool> predicate)
    {
        return RowsUsed(options, predicate);
    }

    internal XlsxRangeRows RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, bool> predicate = null)
    {
        var rows = new XlsxRangeRows(Worksheet);
        int rowCount = RowCount(options);

        for (int ro = 1; ro <= rowCount; ro++)
        {
            XlsxRangeRow row = Row(ro);

            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                rows.Add(row);
        }
        return rows;
    }

    IXlsxRangeRows IXlsxRange.RowsUsed(Func<IXlsxRangeRow, bool> predicate)
    {
        return RowsUsed(predicate);
    }

    internal XlsxRangeRows RowsUsed(Func<IXlsxRangeRow, bool> predicate = null)
    {
        return RowsUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeColumns IXlsxRange.ColumnsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, bool> predicate)
    {
        return ColumnsUsed(options, predicate);
    }

    internal virtual XlsxRangeColumns ColumnsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, bool> predicate = null)
    {
        var columns = new XlsxRangeColumns(Worksheet);
        int columnCount = ColumnCount(options);

        for (int co = 1; co <= columnCount; co++)
        {
            XlsxRangeColumn column = Column(co);

            if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
                columns.Add(column);
        }
        return columns;
    }

    IXlsxRangeColumns IXlsxRange.ColumnsUsed(Func<IXlsxRangeColumn, bool> predicate)
    {
        return ColumnsUsed(predicate);
    }

    internal virtual XlsxRangeColumns ColumnsUsed(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        return ColumnsUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    public XlsxRangeRow Row(int row)
    {
        if (row <= 0 || row > XlsxHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
            throw new ArgumentOutOfRangeException(nameof(row), string.Format("Row number must be between 1 and {0}", XlsxHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1));

        var firstCellAddress = new XlsxAddress(Worksheet,
                                             RangeAddress.FirstAddress.RowNumber + row - 1,
                                             RangeAddress.FirstAddress.ColumnNumber,
                                             false,
                                             false);

        var lastCellAddress = new XlsxAddress(Worksheet,
                                            RangeAddress.FirstAddress.RowNumber + row - 1,
                                            RangeAddress.LastAddress.ColumnNumber,
                                            false,
                                            false);
        return Worksheet.RangeRow(new XlsxRangeAddress(firstCellAddress, lastCellAddress));
    }

    public virtual XlsxRangeColumn Column(int columnNumber)
    {
        if (columnNumber <= 0 || columnNumber > XlsxHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1)
            throw new ArgumentOutOfRangeException(nameof(columnNumber), string.Format("Column number must be between 1 and {0}", XlsxHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1));

        var firstCellAddress = new XlsxAddress(Worksheet,
                                             RangeAddress.FirstAddress.RowNumber,
                                             RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
                                             false,
                                             false);
        var lastCellAddress = new XlsxAddress(Worksheet,
                                            RangeAddress.LastAddress.RowNumber,
                                            RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
                                            false,
                                            false);
        return Worksheet.RangeColumn(new XlsxRangeAddress(firstCellAddress, lastCellAddress));
    }

    public virtual XlsxRangeColumn Column(string columnLetter)
    {
        return Column(XlsxHelper.GetColumnNumberFromLetter(columnLetter));
    }

    internal IEnumerable<XlsxRange> Split(IXlsxRangeAddress anotherRange, bool includeIntersection)
    {
        if (!RangeAddress.Intersects(anotherRange))
        {
            yield return this;
            yield break;
        }

        var thisRow1 = RangeAddress.FirstAddress.RowNumber;
        var thisRow2 = RangeAddress.LastAddress.RowNumber;
        var thisColumn1 = RangeAddress.FirstAddress.ColumnNumber;
        var thisColumn2 = RangeAddress.LastAddress.ColumnNumber;

        var otherRow1 = Math.Min(Math.Max(thisRow1, anotherRange.FirstAddress.RowNumber), thisRow2 + 1);
        var otherRow2 = Math.Max(Math.Min(thisRow2, anotherRange.LastAddress.RowNumber), thisRow1 - 1);
        var otherColumn1 = Math.Min(Math.Max(thisColumn1, anotherRange.FirstAddress.ColumnNumber), thisColumn2 + 1);
        var otherColumn2 = Math.Max(Math.Min(thisColumn2, anotherRange.LastAddress.ColumnNumber), thisColumn1 - 1);

        XlsxRangeAddress[] candidates =
        [
            // to the top of the intersection
            new XlsxRangeAddress(
                new XlsxAddress(thisRow1,thisColumn1, false, false),
                new XlsxAddress(otherRow1 - 1, thisColumn2, false, false)),

            // to the left of the intersection
            new XlsxRangeAddress(
                new XlsxAddress(otherRow1,thisColumn1, false, false),
                new XlsxAddress(otherRow2, otherColumn1 - 1, false, false)),

            includeIntersection
                ? new XlsxRangeAddress(
                    new XlsxAddress(otherRow1, otherColumn1, false, false),
                    new XlsxAddress(otherRow2, otherColumn2, false, false))
                : XlsxRangeAddress.Invalid,

            // to the right of the intersection
            new XlsxRangeAddress(
                new XlsxAddress(otherRow1,otherColumn2 + 1, false, false),
                new XlsxAddress(otherRow2, thisColumn2, false, false)),

            // to the bottom of the intersection
            new XlsxRangeAddress(
                new XlsxAddress(otherRow2 + 1,thisColumn1, false, false),
                new XlsxAddress(thisRow2, thisColumn2, false, false)),
        ];

        foreach (XlsxRangeAddress rangeAddress in candidates.Where(c => c.IsValid && c.IsNormalized))
        {
            yield return Worksheet.Range(rangeAddress);
        }
    }

    private void TransposeRange(int squareSide)
    {
        var rowOffset = RangeAddress.FirstAddress.RowNumber - 1;
        var colOffset = RangeAddress.FirstAddress.ColumnNumber - 1;
        for (var row = 1; row <= squareSide; ++row)
        {
            for (var col = row + 1; col <= squareSide; ++col)
            {
                var oldAddress = new Point(row + rowOffset, col + colOffset);
                var newAddress = new Point(col + colOffset, row + rowOffset);
                Worksheet.Internals.CellsCollection.SwapCellsContent(oldAddress, newAddress);
            }
        }
    }

    private void TransposeMerged(int squareSide)
    {
        XlsxRange rngToTranspose = Worksheet.Range(
            RangeAddress.FirstAddress.RowNumber,
            RangeAddress.FirstAddress.ColumnNumber,
            RangeAddress.FirstAddress.RowNumber + squareSide - 1,
            RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

        foreach (XlsxRange merge in Worksheet.Internals.MergedRanges.Where<XlsxRange>(Contains))
        {
            merge.RangeAddress = new XlsxRangeAddress(
                merge.RangeAddress.FirstAddress,
                rngToTranspose.Cell(merge.ColumnCount(), merge.RowCount()).Address);
        }
    }

    private void MoveOrClearForTranspose(XlsxTransposeOptions transposeOption, int rowCount, int columnCount)
    {
        if (transposeOption == XlsxTransposeOptions.MoveCells)
        {
            if (rowCount > columnCount)
                InsertColumnsAfter(false, rowCount - columnCount, false);
            else if (columnCount > rowCount)
                InsertRowsBelow(false, columnCount - rowCount, false);
        }
        else
        {
            if (rowCount > columnCount)
            {
                int toMove = rowCount - columnCount;
                XlsxRange rngToClear = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + 1,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + toMove);
                rngToClear.Clear();
            }
            else if (columnCount > rowCount)
            {
                int toMove = columnCount - rowCount;
                XlsxRange rngToClear = Worksheet.Range(
                    RangeAddress.LastAddress.RowNumber + 1,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber + toMove,
                    RangeAddress.LastAddress.ColumnNumber);
                rngToClear.Clear();
            }
        }
    }

    public override bool Equals(object obj)
    {
        if (obj is not XlsxRange other)
            return false;
        return RangeAddress.Equals(other.RangeAddress)
               && Worksheet.Equals(other.Worksheet);
    }

    public override int GetHashCode()
    {
        return RangeAddress.GetHashCode()
               ^ Worksheet.GetHashCode();
    }

    public new IXlsxRange Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXlsxRangeColumn FindColumn(Func<IXlsxRangeColumn, bool> predicate)
    {
        int columnCount = ColumnCount();
        for (int c = 1; c <= columnCount; c++)
        {
            XlsxRangeColumn column = Column(c);
            if (predicate == null || predicate(column))
                return column;
        }
        return null;
    }

    public IXlsxRangeRow FindRow(Func<IXlsxRangeRow, bool> predicate)
    {
        int rowCount = RowCount();
        for (int r = 1; r <= rowCount; r++)
        {
            XlsxRangeRow row = Row(r);
            if (predicate(row))
                return row;
        }
        return null;
    }

    public override string ToString()
    {
        if (IsEntireSheet())
        {
            return Worksheet.Name;
        }
        else if (IsEntireRow())
        {
            return string.Concat(
                Worksheet.Name.EscapeSheetName(),
                '!',
                RangeAddress.FirstAddress.RowNumber,
                ':',
                RangeAddress.LastAddress.RowNumber);
        }
        else if (IsEntireColumn())
        {
            return string.Concat(
                Worksheet.Name.EscapeSheetName(),
                '!',
                RangeAddress.FirstAddress.ColumnLetter,
                ':',
                RangeAddress.LastAddress.ColumnLetter);
        }
        else
            return base.ToString();
    }
}
