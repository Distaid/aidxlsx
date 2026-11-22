#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.Tables
{
    using System.Linq;

    internal class XlsxTableRange : XlsxRange, IXlsxTableRange
    {
        private readonly XlsxTable _table;
        private readonly XlsxRange _range;

        public XlsxTableRange(XlsxRange range, XlsxTable table)
            : base(range.RangeAddress, range.Style)
        {
            _table = table;
            _range = range;
        }

        IXlsxTableRow IXlsxTableRange.FirstRow(Func<IXlsxTableRow, Boolean> predicate)
        {
            return FirstRow(predicate);
        }

        public XlsxTableRow FirstRow(Func<IXlsxTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XlsxTableRow(this, (_range.FirstRow()));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = new XlsxTableRow(this, (_range.Row(ro)));
                if (predicate(row)) return row;
            }

            return null;
        }

        IXlsxTableRow IXlsxTableRange.FirstRowUsed(Func<IXlsxTableRow, Boolean> predicate)
        {
            return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
        }

        public XlsxTableRow FirstRowUsed(Func<IXlsxTableRow, Boolean> predicate = null)
        {
            return FirstRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
        }

        IXlsxTableRow IXlsxTableRange.FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, Boolean> predicate)
        {
            return FirstRowUsed(options, predicate);
        }

        internal XlsxTableRow FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XlsxTableRow(this, (_range.FirstRowUsed(options)));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = new XlsxTableRow(this, (_range.Row(ro)));

                if (!row.IsEmpty(options) && predicate(row))
                    return row;
            }

            return null;
        }

        IXlsxTableRow IXlsxTableRange.LastRow(Func<IXlsxTableRow, Boolean> predicate)
        {
            return LastRow(predicate);
        }

        public XlsxTableRow LastRow(Func<IXlsxTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XlsxTableRow(this, (_range.LastRow()));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = rowCount; ro >= 1; ro--)
            {
                var row = new XlsxTableRow(this, (_range.Row(ro)));
                if (predicate(row)) return row;
            }
            return null;
        }

        IXlsxTableRow IXlsxTableRange.LastRowUsed(Func<IXlsxTableRow, Boolean> predicate)
        {
            return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
        }

        public XlsxTableRow LastRowUsed(Func<IXlsxTableRow, Boolean> predicate = null)
        {
            return LastRowUsed(XlsxCellsUsedOptions.AllContents, predicate);
        }

        IXlsxTableRow IXlsxTableRange.LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, Boolean> predicate)
        {
            return LastRowUsed(options, predicate);
        }

        internal XlsxTableRow LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XlsxTableRow(this, (_range.LastRowUsed(options)));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = rowCount; ro >= 1; ro--)
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
                    String.Format("Row number must be between 1 and {0}", XlsxHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
                );
            }

            return new XlsxTableRow(this, base.Row(row));
        }

        public IXlsxTableRows Rows(Func<IXlsxTableRow, Boolean> predicate = null)
        {
            var retVal = new XlsxTableRows(Worksheet);
            Int32 rowCount = _range.RowCount();

            for (int r = 1; r <= rowCount; r++)
            {
                var row = Row(r);
                if (predicate == null || predicate(row))
                    retVal.Add(row);
            }
            return retVal;
        }

        public new IXlsxTableRows Rows(int firstRow, int lastRow)
        {
            var retVal = new XlsxTableRows(Worksheet);

            for (int rowNumber = firstRow; rowNumber <= lastRow; rowNumber++)
                retVal.Add(Row(rowNumber));

            return retVal;
        }

        public new IXlsxTableRows Rows(string rows)
        {
            var retVal = new XlsxTableRows(Worksheet);
            var rowPairs = rows.Split(',');
            foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
            {
                String firstRow;
                String lastRow;
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
                foreach (IXlsxTableRow row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                    retVal.Add(row);
            }
            return retVal;
        }

        IXlsxTableRows IXlsxTableRange.RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, Boolean> predicate)
        {
            return RowsUsed(options, predicate);
        }

        internal XlsxTableRows RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, Boolean> predicate = null)
        {
            var rows = new XlsxTableRows(Worksheet);
            Int32 rowCount = RowCount();

            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                    rows.Add(row);
            }
            return rows;
        }

        IXlsxTableRows IXlsxTableRange.RowsUsed(Func<IXlsxTableRow, Boolean> predicate)
        {
            return RowsUsed(predicate);
        }

        public IXlsxTableRows RowsUsed(Func<IXlsxTableRow, Boolean> predicate = null)
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

        public new IXlsxRangeColumn Column(String column)
        {
            if (XlsxHelper.IsValidColumn(column))
            {
                Int32 coNum = XlsxHelper.GetColumnNumberFromLetter(column);
                return coNum > ColumnCount() ? Column(_table.GetFieldIndex(column) + 1) : Column(coNum);
            }

            return Column(_table.GetFieldIndex(column) + 1);
        }
    }
}
