#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Ranges.Sort;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.AutoFilters
{
    using System.Collections.Generic;

    internal class XlsxAutoFilter : IXlsxAutoFilter
    {
        /// <summary>
        /// Key is column number.
        /// </summary>
        private readonly Dictionary<Int32, XlsxFilterColumn> _columns = new();

        internal IReadOnlyDictionary<Int32, XlsxFilterColumn> Columns => _columns;

        #region IXlsxAutoFilter Members

        public IEnumerable<IXlsxRangeRow> HiddenRows => Range.Rows(r => r.WorksheetRow().IsHidden);

        public Boolean IsEnabled { get; set; }

        public IXlsxRange Range { get; set; }

        public Int32 SortColumn { get; set; }

        public Boolean Sorted { get; set; }

        public XlsxSortOrder SortOrder { get; set; }

        public IEnumerable<IXlsxRangeRow> VisibleRows => Range.Rows(r => !r.WorksheetRow().IsHidden);

        IXlsxAutoFilter IXlsxAutoFilter.Clear() => Clear();

        IXlsxFilterColumn IXlsxAutoFilter.Column(String columnLetter) => Column(columnLetter);

        IXlsxFilterColumn IXlsxAutoFilter.Column(Int32 columnNumber) => Column(columnNumber);

        IXlsxAutoFilter IXlsxAutoFilter.Sort(Int32 columnToSortBy, XlsxSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        public IXlsxAutoFilter Reapply()
        {
            // Recalculate shown / hidden rows
            var rows = Range.Rows(2, Range.RowCount());
            rows.ForEach(row =>
                row.WorksheetRow().Unhide()
            );

            foreach (var filterColumn in _columns.Values)
                filterColumn.Refresh();

            foreach (IXlsxRangeRow row in rows)
            {
                var rowMatch = true;

                foreach (var (columnIndex, column) in _columns)
                {
                    var cell = row.Cell(columnIndex);
                    var columnFilterMatch = column.Check(cell);
                    rowMatch &= columnFilterMatch;

                    if (!rowMatch) break;
                }

                if (!rowMatch) row.WorksheetRow().Hide();
            }

            return this;
        }

        #endregion IXlsxAutoFilter Members

        internal XlsxFilterColumn Column(String columnLetter)
        {
            var columnNumber = XlsxHelper.GetColumnNumberFromLetter(columnLetter);
            if (columnNumber < 1 || columnNumber > XlsxHelper.MaxColumnNumber)
                throw new ArgumentOutOfRangeException(nameof(columnLetter), "Column '" + columnLetter + "' is outside the allowed column range.");

            return Column(columnNumber);
        }

        internal XlsxFilterColumn Column(Int32 columnNumber)
        {
            if (columnNumber < 1 || columnNumber > XlsxHelper.MaxColumnNumber)
                throw new ArgumentOutOfRangeException(nameof(columnNumber), "Column " + columnNumber + " is outside the allowed column range.");

            if (!_columns.TryGetValue(columnNumber, out XlsxFilterColumn filterColumn))
            {
                filterColumn = new XlsxFilterColumn(this, columnNumber);
                _columns.Add(columnNumber, filterColumn);
            }

            return filterColumn;
        }

        internal XlsxAutoFilter Clear()
        {
            if (!IsEnabled) return this;

            IsEnabled = false;
            foreach (var filterColumn in _columns.Values)
                filterColumn.Clear(false);

            foreach (IXlsxRangeRow row in Range.Rows().Where(r => r.RowNumber() > 1))
                row.WorksheetRow().Unhide();
            return this;
        }

        internal XlsxAutoFilter Set(IXlsxRangeBase range)
        {
            var firstOverlappingTable = range.Worksheet.Tables.FirstOrDefault(t => t.RangeUsed().Intersects(range));
            if (firstOverlappingTable != null)
                throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'");

            Range = range.AsRange();
            IsEnabled = true;
            return this;
        }

        internal XlsxAutoFilter Sort(Int32 columnToSortBy, XlsxSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            if (!IsEnabled)
                throw new InvalidOperationException("Filter has not been enabled.");

            Range.Range(Range.FirstCell().CellBelow(), Range.LastCell()).Sort(columnToSortBy, sortOrder, matchCase,
                                                                              ignoreBlanks);

            Sorted = true;
            SortOrder = sortOrder;
            SortColumn = columnToSortBy;

            Reapply();

            return this;
        }
    }
}
