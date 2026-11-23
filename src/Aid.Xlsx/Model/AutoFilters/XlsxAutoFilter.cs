#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Tables;

namespace Aid.Xlsx.Model.AutoFilters;

using System.Collections.Generic;

internal class XlsxAutoFilter : IXlsxAutoFilter
{
    /// <summary>
    /// Key is column number.
    /// </summary>
    private readonly Dictionary<int, XlsxFilterColumn> _columns = [];

    internal IReadOnlyDictionary<int, XlsxFilterColumn> Columns => _columns;

    #region IXlsxAutoFilter Members

    public IEnumerable<IXlsxRangeRow> HiddenRows => Range.Rows(r => r.WorksheetRow().IsHidden);

    public bool IsEnabled { get; set; }

    public IXlsxRange Range { get; set; }

    public int SortColumn { get; set; }

    public bool Sorted { get; set; }

    public XlsxSortOrder SortOrder { get; set; }

    public IEnumerable<IXlsxRangeRow> VisibleRows => Range.Rows(r => !r.WorksheetRow().IsHidden);

    IXlsxAutoFilter IXlsxAutoFilter.Clear() => Clear();

    IXlsxFilterColumn IXlsxAutoFilter.Column(string columnLetter) => Column(columnLetter);

    IXlsxFilterColumn IXlsxAutoFilter.Column(int columnNumber) => Column(columnNumber);

    IXlsxAutoFilter IXlsxAutoFilter.Sort(int columnToSortBy, XlsxSortOrder sortOrder, bool matchCase, bool ignoreBlanks)
    {
        return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
    }

    public IXlsxAutoFilter Reapply()
    {
        // Recalculate shown / hidden rows
        IXlsxRangeRows rows = Range.Rows(2, Range.RowCount());
        rows.ForEach(row =>
            row.WorksheetRow().Unhide()
        );

        foreach (XlsxFilterColumn filterColumn in _columns.Values)
            filterColumn.Refresh();

        foreach (IXlsxRangeRow row in rows)
        {
            var rowMatch = true;

            foreach ((var columnIndex, XlsxFilterColumn column) in _columns)
            {
                IXlsxCell cell = row.Cell(columnIndex);
                var columnFilterMatch = column.Check(cell);
                rowMatch &= columnFilterMatch;

                if (!rowMatch) break;
            }

            if (!rowMatch) row.WorksheetRow().Hide();
        }

        return this;
    }

    #endregion IXlsxAutoFilter Members

    internal XlsxFilterColumn Column(string columnLetter)
    {
        var columnNumber = XlsxHelper.GetColumnNumberFromLetter(columnLetter);
        if (columnNumber < 1 || columnNumber > XlsxHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnLetter), "Column '" + columnLetter + "' is outside the allowed column range.");

        return Column(columnNumber);
    }

    internal XlsxFilterColumn Column(int columnNumber)
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
        foreach (XlsxFilterColumn filterColumn in _columns.Values)
            filterColumn.Clear(false);

        foreach (IXlsxRangeRow row in Range.Rows().Where(r => r.RowNumber() > 1))
            row.WorksheetRow().Unhide();
        return this;
    }

    internal XlsxAutoFilter Set(IXlsxRangeBase range)
    {
        IXlsxTable firstOverlappingTable = range.Worksheet.Tables.FirstOrDefault(t => t.RangeUsed().Intersects(range));
        if (firstOverlappingTable != null)
            throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'");

        Range = range.AsRange();
        IsEnabled = true;
        return this;
    }

    internal XlsxAutoFilter Sort(int columnToSortBy, XlsxSortOrder sortOrder, bool matchCase, bool ignoreBlanks)
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
