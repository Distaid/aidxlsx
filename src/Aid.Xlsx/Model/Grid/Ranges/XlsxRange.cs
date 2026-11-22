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

internal class XlsxRange : XlsxRangeBase, IXlsxRange
{
    public XlsxRange(XlsxRangeAddress rangeAddress, IXlsxStyle defaultStyle)
        : base(rangeAddress)
    {
    }

    internal override XlsxCellFormat Format => XlsxCellFormat.ForRange(Worksheet, RangeAddress);

    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.Range; }
    }

    #region IXlsxRange Members

    IXlsxRangeRow IXlsxRange.Row(Int32 row)
    {
        return Row(row);
    }

    IXlsxRangeColumn IXlsxRange.Column(Int32 columnNumber)
    {
        return Column(columnNumber);
    }

    IXlsxRangeColumn IXlsxRange.Column(String columnLetter)
    {
        return Column(columnLetter);
    }

    public virtual IXlsxRangeColumns Columns(Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        var retVal = new XlsxRangeColumns(Worksheet);
        Int32 columnCount = ColumnCount();
        for (Int32 c = 1; c <= columnCount; c++)
        {
            var column = Column(c);
            if (predicate == null || predicate(column))
                retVal.Add(column);
        }
        return retVal;
    }

    public virtual IXlsxRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
    {
        var retVal = new XlsxRangeColumns(Worksheet);

        for (int co = firstColumn; co <= lastColumn; co++)
            retVal.Add(Column(co));
        return retVal;
    }

    public virtual IXlsxRangeColumns Columns(String firstColumn, String lastColumn)
    {
        return Columns(XlsxHelper.GetColumnNumberFromLetter(firstColumn),
                       XlsxHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public virtual IXlsxRangeColumns Columns(String columns)
    {
        var retVal = new XlsxRangeColumns(Worksheet);
        var columnPairs = columns.Split(',');
        foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
        {
            String firstColumn;
            String lastColumn;
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

            if (Int32.TryParse(firstColumn, out Int32 tmp))
            {
                foreach (IXlsxRangeColumn col in Columns(Int32.Parse(firstColumn), Int32.Parse(lastColumn)))
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

    public IXlsxRangeRows Rows(Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        var retVal = new XlsxRangeRows(Worksheet);
        Int32 rowCount = RowCount();
        for (Int32 r = 1; r <= rowCount; r++)
        {
            var row = Row(r);
            if (predicate == null || predicate(row))
                retVal.Add(Row(r));
        }
        return retVal;
    }

    public IXlsxRangeRows Rows(Int32 firstRow, Int32 lastRow)
    {
        var retVal = new XlsxRangeRows(Worksheet);

        for (int ro = firstRow; ro <= lastRow; ro++)
            retVal.Add(Row(ro));
        return retVal;
    }

    public IXlsxRangeRows Rows(String rows)
    {
        var retVal = new XlsxRangeRows(Worksheet);
        var rowPairs = rows.Split(',');
        foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
        {
            String firstRow;
            String lastRow;
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
            foreach (IXlsxRangeRow row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                retVal.Add(row);
        }
        return retVal;
    }

    public void Transpose(XlsxTransposeOptions transposeOption)
    {
        int rowCount = RowCount();
        int columnCount = ColumnCount();
        int squareSide = rowCount > columnCount ? rowCount : columnCount;

        var firstCell = FirstCell();

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
            var rng = Worksheet.Range(
                RangeAddress.LastAddress.RowNumber + 1,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber + (rowCount - columnCount),
                RangeAddress.LastAddress.ColumnNumber);
            rng.Delete(XlsxShiftDeletedCells.ShiftCellsUp);
        }
        else if (columnCount > rowCount)
        {
            var rng = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber + 1,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber + (columnCount - rowCount));
            rng.Delete(XlsxShiftDeletedCells.ShiftCellsLeft);
        }

        foreach (IXlsxCell c in Range(1, 1, columnCount, rowCount).Cells())
        {
            var border = ((XlsxCell)c).GetFormat().Border;
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

    public IXlsxTable AsTable(String name)
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

    IXlsxTable IXlsxRange.CreateTable(String name)
    {
        return CreateTable(name);
    }

    public XlsxTable CreateTable(String name)
    {
        return (XlsxTable)Worksheet.Table(this, name, true, true);
    }

    public IXlsxTable CreateTable(String name, Boolean setAutofilter)
    {
        return Worksheet.Table(this, name, true, setAutofilter);
    }

    public IXlsxRange CopyTo(IXlsxCell target)
    {
        base.CopyTo((XlsxCell)target);

        Int32 lastRowNumber = target.Address.RowNumber + RowCount() - 1;
        if (lastRowNumber > XlsxHelper.MaxRowNumber)
            lastRowNumber = XlsxHelper.MaxRowNumber;
        Int32 lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
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

        Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
        if (lastRowNumber > XlsxHelper.MaxRowNumber)
            lastRowNumber = XlsxHelper.MaxRowNumber;
        Int32 lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
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

    public new IXlsxRange Sort(String columnsToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
    {
        return base.Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    public new IXlsxRange Sort(Int32 columnToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
    {
        return base.Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    public new IXlsxRange SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
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

    IXlsxRangeColumn IXlsxRange.FirstColumn(Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return FirstColumn(predicate);
    }

    internal XlsxRangeColumn FirstColumn(Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        if (predicate == null)
            return Column(1);

        Int32 columnCount = ColumnCount();
        for (Int32 c = 1; c <= columnCount; c++)
        {
            var column = Column(c);
            if (predicate(column)) return column;
        }

        return null;
    }

    IXlsxRangeColumn IXlsxRange.LastColumn(Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return LastColumn(predicate);
    }

    internal XlsxRangeColumn LastColumn(Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        Int32 columnCount = ColumnCount();
        if (predicate == null)
            return Column(columnCount);

        for (Int32 c = columnCount; c >= 1; c--)
        {
            var column = Column(c);
            if (predicate(column)) return column;
        }

        return null;
    }

    IXlsxRangeColumn IXlsxRange.FirstColumnUsed(Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return FirstColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeColumn FirstColumnUsed(Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        return FirstColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeColumn IXlsxRange.FirstColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return FirstColumnUsed(options, predicate);
    }

    internal XlsxRangeColumn FirstColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            Int32 firstColumnUsed = Worksheet.Internals.CellsCollection.FirstColumnUsed(
                Area.FromRangeAddress(RangeAddress),
                options);

            return firstColumnUsed == 0 ? null : Column(firstColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        Int32 columnCount = ColumnCount();
        for (Int32 co = 1; co <= columnCount; co++)
        {
            var column = Column(co);

            if (!column.IsEmpty(options) && predicate(column))
                return column;
        }
        return null;
    }

    IXlsxRangeColumn IXlsxRange.LastColumnUsed(Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return LastColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeColumn LastColumnUsed(Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        return LastColumnUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeColumn IXlsxRange.LastColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return LastColumnUsed(options, predicate);
    }

    internal XlsxRangeColumn LastColumnUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            Int32 lastColumnUsed = Worksheet.Internals.CellsCollection.LastColumnUsed(
                Area.FromRangeAddress(RangeAddress),
                options);

            return lastColumnUsed == 0 ? null : Column(lastColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        Int32 columnCount = ColumnCount();
        for (Int32 co = columnCount; co >= 1; co--)
        {
            var column = Column(co);

            if (!column.IsEmpty(options) && predicate(column))
                return column;
        }
        return null;
    }

    IXlsxRangeRow IXlsxRange.FirstRow(Func<IXlsxRangeRow, Boolean> predicate)
    {
        return FirstRow(predicate);
    }

    public XlsxRangeRow FirstRow(Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        if (predicate == null)
            return Row(1);

        Int32 rowCount = RowCount();
        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            var row = Row(ro);
            if (predicate(row)) return row;
        }

        return null;
    }

    IXlsxRangeRow IXlsxRange.LastRow(Func<IXlsxRangeRow, Boolean> predicate)
    {
        return LastRow(predicate);
    }

    public XlsxRangeRow LastRow(Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        Int32 rowCount = RowCount();
        if (predicate == null)
            return Row(rowCount);

        for (Int32 ro = rowCount; ro >= 1; ro--)
        {
            var row = Row(ro);
            if (predicate(row)) return row;
        }

        return null;
    }

    IXlsxRangeRow IXlsxRange.FirstRowUsed(Func<IXlsxRangeRow, Boolean> predicate)
    {
        return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeRow FirstRowUsed(Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeRow IXlsxRange.FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, Boolean> predicate)
    {
        return FirstRowUsed(options, predicate);
    }

    internal XlsxRangeRow FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            Int32 rowFromCells = Worksheet.Internals.CellsCollection.FirstRowUsed(
                Area.FromRangeAddress(RangeAddress), options);

            return rowFromCells == 0 ? null : Row(rowFromCells - RangeAddress.FirstAddress.RowNumber + 1);
        }

        Int32 rowCount = RowCount();
        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            var row = Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }
        return null;
    }

    IXlsxRangeRow IXlsxRange.LastRowUsed(Func<IXlsxRangeRow, Boolean> predicate)
    {
        return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxRangeRow LastRowUsed(Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeRow IXlsxRange.LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, Boolean> predicate)
    {
        return LastRowUsed(options, predicate);
    }

    internal XlsxRangeRow LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            Int32 lastRowUsed = Worksheet.Internals.CellsCollection.LastRowUsed(
                Area.FromRangeAddress(RangeAddress), options);

            return lastRowUsed == 0 ? null : Row(lastRowUsed - RangeAddress.FirstAddress.RowNumber + 1);
        }

        Int32 rowCount = RowCount();
        for (Int32 ro = rowCount; ro >= 1; ro--)
        {
            var row = Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }
        return null;
    }

    IXlsxRangeRows IXlsxRange.RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, Boolean> predicate)
    {
        return RowsUsed(options, predicate);
    }

    internal XlsxRangeRows RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        XlsxRangeRows rows = new XlsxRangeRows(Worksheet);
        Int32 rowCount = RowCount(options);

        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            var row = Row(ro);

            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                rows.Add(row);
        }
        return rows;
    }

    IXlsxRangeRows IXlsxRange.RowsUsed(Func<IXlsxRangeRow, Boolean> predicate)
    {
        return RowsUsed(predicate);
    }

    internal XlsxRangeRows RowsUsed(Func<IXlsxRangeRow, Boolean> predicate = null)
    {
        return RowsUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxRangeColumns IXlsxRange.ColumnsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return ColumnsUsed(options, predicate);
    }

    internal virtual XlsxRangeColumns ColumnsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        XlsxRangeColumns columns = new XlsxRangeColumns(Worksheet);
        Int32 columnCount = ColumnCount(options);

        for (Int32 co = 1; co <= columnCount; co++)
        {
            var column = Column(co);

            if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
                columns.Add(column);
        }
        return columns;
    }

    IXlsxRangeColumns IXlsxRange.ColumnsUsed(Func<IXlsxRangeColumn, Boolean> predicate)
    {
        return ColumnsUsed(predicate);
    }

    internal virtual XlsxRangeColumns ColumnsUsed(Func<IXlsxRangeColumn, Boolean> predicate = null)
    {
        return ColumnsUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    public XlsxRangeRow Row(Int32 row)
    {
        if (row <= 0 || row > XlsxHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
            throw new ArgumentOutOfRangeException(nameof(row), String.Format("Row number must be between 1 and {0}", XlsxHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1));

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

    public virtual XlsxRangeColumn Column(Int32 columnNumber)
    {
        if (columnNumber <= 0 || columnNumber > XlsxHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1)
            throw new ArgumentOutOfRangeException(nameof(columnNumber), String.Format("Column number must be between 1 and {0}", XlsxHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1));

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

    public virtual XlsxRangeColumn Column(String columnLetter)
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

        var candidates = new[]
        {
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
        };

        foreach (var rangeAddress in candidates.Where(c => c.IsValid && c.IsNormalized))
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

    private void TransposeMerged(Int32 squareSide)
    {
        var rngToTranspose = Worksheet.Range(
            RangeAddress.FirstAddress.RowNumber,
            RangeAddress.FirstAddress.ColumnNumber,
            RangeAddress.FirstAddress.RowNumber + squareSide - 1,
            RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

        foreach (var merge in Worksheet.Internals.MergedRanges.Where<XlsxRange>(Contains))
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
                var rngToClear = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + 1,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + toMove);
                rngToClear.Clear();
            }
            else if (columnCount > rowCount)
            {
                int toMove = columnCount - rowCount;
                var rngToClear = Worksheet.Range(
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
        var other = obj as XlsxRange;
        if (other == null)
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
        Int32 columnCount = ColumnCount();
        for (Int32 c = 1; c <= columnCount; c++)
        {
            var column = Column(c);
            if (predicate == null || predicate(column))
                return column;
        }
        return null;
    }

    public IXlsxRangeRow FindRow(Func<IXlsxRangeRow, bool> predicate)
    {
        Int32 rowCount = RowCount();
        for (Int32 r = 1; r <= rowCount; r++)
        {
            var row = Row(r);
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
            return String.Concat(
                Worksheet.Name.EscapeSheetName(),
                '!',
                RangeAddress.FirstAddress.RowNumber,
                ':',
                RangeAddress.LastAddress.RowNumber);
        }
        else if (IsEntireColumn())
        {
            return String.Concat(
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
