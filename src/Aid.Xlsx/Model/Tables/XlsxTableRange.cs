#nullable disable

using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.Tables;

using System.Linq;

internal class XlsxTableRange(XlsxRange range, XlsxTable table) : XlsxRange(range.RangeAddress), IXlsxTableRange
{
    private readonly XlsxTable _table = table;
    private readonly XlsxRange _range = range;

    IXlsxTableRow IXlsxTableRange.FirstRow(Func<IXlsxTableRow, bool> predicate)
    {
        return FirstRow(predicate);
    }

    public XlsxTableRow FirstRow(Func<IXlsxTableRow, bool> predicate = null)
    {
        if (predicate == null)
            return new XlsxTableRow(this, (_range.FirstRow()));

        var rowCount = _range.RowCount();

        for (var ro = 1; ro <= rowCount; ro++)
        {
            var row = new XlsxTableRow(this, (_range.Row(ro)));
            if (predicate(row)) return row;
        }

        return null;
    }

    IXlsxTableRow IXlsxTableRange.FirstRowUsed(Func<IXlsxTableRow, bool> predicate)
    {
        return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    public XlsxTableRow FirstRowUsed(Func<IXlsxTableRow, bool> predicate = null)
    {
        return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxTableRow IXlsxTableRange.FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate)
    {
        return FirstRowUsed(options, predicate);
    }

    internal XlsxTableRow FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate = null)
    {
        if (predicate == null)
            return new XlsxTableRow(this, (_range.FirstRowUsed(options)));

        var rowCount = _range.RowCount();

        for (var ro = 1; ro <= rowCount; ro++)
        {
            var row = new XlsxTableRow(this, (_range.Row(ro)));

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }

        return null;
    }

    IXlsxTableRow IXlsxTableRange.LastRow(Func<IXlsxTableRow, bool> predicate)
    {
        return LastRow(predicate);
    }

    public XlsxTableRow LastRow(Func<IXlsxTableRow, bool> predicate = null)
    {
        if (predicate == null)
            return new XlsxTableRow(this, (_range.LastRow()));

        var rowCount = _range.RowCount();

        for (var ro = rowCount; ro >= 1; ro--)
        {
            var row = new XlsxTableRow(this, (_range.Row(ro)));
            if (predicate(row)) return row;
        }
        return null;
    }

    IXlsxTableRow IXlsxTableRange.LastRowUsed(Func<IXlsxTableRow, bool> predicate)
    {
        return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    public XlsxTableRow LastRowUsed(Func<IXlsxTableRow, bool> predicate = null)
    {
        return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxTableRow IXlsxTableRange.LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate)
    {
        return LastRowUsed(options, predicate);
    }

    internal XlsxTableRow LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate = null)
    {
        if (predicate == null)
            return new XlsxTableRow(this, (_range.LastRowUsed(options)));

        var rowCount = _range.RowCount();

        for (var ro = rowCount; ro >= 1; ro--)
        {
            var row = new XlsxTableRow(this, (_range.Row(ro)));

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }

        return null;
    }

    IXlsxTableRow IXlsxTableRange.Row(int row)
    {
        return Row(row);
    }

    public new XlsxTableRow Row(int row)
    {
        if (row <= 0 || row > XlsxHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(row),
                string.Format("Row number must be between 1 and {0}", XlsxHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
            );
        }

        return new XlsxTableRow(this, base.Row(row));
    }

    public IXlsxTableRows Rows(Func<IXlsxTableRow, bool> predicate = null)
    {
        var retVal = new XlsxTableRows(Worksheet);
        var rowCount = _range.RowCount();

        for (var r = 1; r <= rowCount; r++)
        {
            XlsxTableRow row = Row(r);
            if (predicate == null || predicate(row))
                retVal.Add(row);
        }
        return retVal;
    }

    public new IXlsxTableRows Rows(int firstRow, int lastRow)
    {
        var retVal = new XlsxTableRows(Worksheet);

        for (var rowNumber = firstRow; rowNumber <= lastRow; rowNumber++)
            retVal.Add(Row(rowNumber));

        return retVal;
    }

    public new IXlsxTableRows Rows(string rows)
    {
        var retVal = new XlsxTableRows(Worksheet);
        var rowPairs = rows.Split(',');
        foreach (var tPair in rowPairs.Select(pair => pair.Trim()))
        {
            string firstRow;
            string lastRow;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                var rowRange = XlsxHelper.SplitRange(tPair);

                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = tPair;
                lastRow = tPair;
            }
            foreach (IXlsxTableRow row in Rows(int.Parse(firstRow), int.Parse(lastRow)))
                retVal.Add(row);
        }
        return retVal;
    }

    IXlsxTableRows IXlsxTableRange.RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate)
    {
        return RowsUsed(options, predicate);
    }

    internal XlsxTableRows RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate = null)
    {
        var rows = new XlsxTableRows(Worksheet);
        var rowCount = RowCount();

        for (var ro = 1; ro <= rowCount; ro++)
        {
            XlsxTableRow row = Row(ro);

            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                rows.Add(row);
        }
        return rows;
    }

    IXlsxTableRows IXlsxTableRange.RowsUsed(Func<IXlsxTableRow, bool> predicate)
    {
        return RowsUsed(predicate);
    }

    public IXlsxTableRows RowsUsed(Func<IXlsxTableRow, bool> predicate = null)
    {
        return RowsUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    IXlsxTable IXlsxTableRange.Table { get { return _table; } }

    public XlsxTable Table { get { return _table; } }

    public new IXlsxTableRows InsertRowsAbove(int numberOfRows)
    {
        return XlsxHelper.InsertRowsWithoutEvents(base.InsertRowsAbove, this, numberOfRows, !Table.ShowTotalsRow);
    }

    public new IXlsxTableRows InsertRowsBelow(int numberOfRows)
    {
        return XlsxHelper.InsertRowsWithoutEvents(base.InsertRowsBelow, this, numberOfRows, !Table.ShowTotalsRow);
    }

    public new IXlsxRangeColumn Column(string column)
    {
        if (XlsxHelper.IsValidColumn(column))
        {
            var coNum = XlsxHelper.GetColumnNumberFromLetter(column);
            return coNum > ColumnCount() ? Column(_table.GetFieldIndex(column) + 1) : Column(coNum);
        }

        return Column(_table.GetFieldIndex(column) + 1);
    }
}
