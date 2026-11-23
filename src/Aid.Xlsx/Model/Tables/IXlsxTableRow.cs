#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;

namespace Aid.Xlsx.Model.Tables;

public interface IXlsxTableRow : IXlsxRangeRow
{
    IXlsxCell Field(int index);

    IXlsxCell Field(string name);

    new IXlsxTableRow Sort();

    new IXlsxTableRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

    new IXlsxTableRow RowAbove();

    new IXlsxTableRow RowAbove(int step);

    new IXlsxTableRow RowBelow();

    new IXlsxTableRow RowBelow(int step);

    /// <summary>
    /// Clears the contents of this row.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    new IXlsxTableRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    new IXlsxTableRows InsertRowsAbove(int numberOfRows);

    new IXlsxTableRows InsertRowsBelow(int numberOfRows);
}
