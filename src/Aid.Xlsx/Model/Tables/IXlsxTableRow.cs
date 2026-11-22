#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;

namespace Aid.Xlsx.Model.Tables;

public interface IXlsxTableRow : IXlsxRangeRow
{
    IXlsxCell Field(Int32 index);

    IXlsxCell Field(String name);

    new IXlsxTableRow Sort();

    new IXlsxTableRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

    new IXlsxTableRow RowAbove();

    new IXlsxTableRow RowAbove(Int32 step);

    new IXlsxTableRow RowBelow();

    new IXlsxTableRow RowBelow(Int32 step);

    /// <summary>
    /// Clears the contents of this row.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    new IXlsxTableRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    new IXlsxTableRows InsertRowsAbove(int numberOfRows);

    new IXlsxTableRows InsertRowsBelow(int numberOfRows);
}
