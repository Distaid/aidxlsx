#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.Tables;

public interface IXlsxTableRange : IXlsxRange
{
    IXlsxTable Table { get; }

    IXlsxTableRow FirstRow(Func<IXlsxTableRow, bool> predicate = null);

    IXlsxTableRow FirstRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate = null);

    IXlsxTableRow FirstRowUsed(Func<IXlsxTableRow, bool> predicate = null);

    new IXlsxTableRows InsertRowsAbove(int numberOfRows);

    new IXlsxTableRows InsertRowsBelow(int numberOfRows);

    IXlsxTableRow LastRow(Func<IXlsxTableRow, bool> predicate = null);

    IXlsxTableRow LastRowUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate = null);

    IXlsxTableRow LastRowUsed(Func<IXlsxTableRow, bool> predicate = null);

    /// <summary>
    /// Rows the specified row.
    /// </summary>
    /// <param name="row">1-based row number relative to the first row of this range.</param>
    new IXlsxTableRow Row(int row);

    IXlsxTableRows Rows(Func<IXlsxTableRow, bool> predicate = null);

    /// <summary>
    /// Returns a subset of the rows
    /// </summary>
    /// <param name="firstRow">The first row to return. 1-based row number relative to the first row of this range.</param>
    /// <param name="lastRow">The last row to return. 1-based row number relative to the first row of this range.</param>
    new IXlsxTableRows Rows(int firstRow, int lastRow);

    new IXlsxTableRows Rows(string rows);

    IXlsxTableRows RowsUsed(XlsxCellsUsedOptions options, Func<IXlsxTableRow, bool> predicate = null);

    IXlsxTableRows RowsUsed(Func<IXlsxTableRow, bool> predicate = null);
}
