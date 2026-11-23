#nullable disable

using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;

namespace Aid.Xlsx.Model.Tables;

internal class XlsxTableRow(XlsxTableRange tableRange, XlsxRangeRow rangeRow) : XlsxRangeRow(rangeRow.RangeAddress), IXlsxTableRow
{
    private readonly XlsxTableRange _tableRange = tableRange;

    #region IXlsxTableRow Members

    public IXlsxCell Field(int index)
    {
        return Cell(index + 1);
    }

    public IXlsxCell Field(string name)
    {
        var fieldIndex = _tableRange.Table.GetFieldIndex(name);
        return Cell(fieldIndex + 1);
    }

    public new IXlsxTableRow Sort()
    {
        return SortLeftToRight();
    }

    public new IXlsxTableRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    #endregion IXlsxTableRow Members

    private XlsxTableRow RowShift(int rowsToShift)
    {
        return _tableRange.Row(RowNumber() - _tableRange.FirstRow().RowNumber() + 1 + rowsToShift);
    }

    #region XlsxTableRow Above

    IXlsxTableRow IXlsxTableRow.RowAbove()
    {
        return RowAbove();
    }

    IXlsxTableRow IXlsxTableRow.RowAbove(int step)
    {
        return RowAbove(step);
    }

    public new XlsxTableRow RowAbove()
    {
        return RowAbove(1);
    }

    public new XlsxTableRow RowAbove(int step)
    {
        return RowShift(step * -1);
    }

    #endregion XlsxTableRow Above

    #region XlsxTableRow Below

    IXlsxTableRow IXlsxTableRow.RowBelow()
    {
        return RowBelow();
    }

    IXlsxTableRow IXlsxTableRow.RowBelow(int step)
    {
        return RowBelow(step);
    }

    public new XlsxTableRow RowBelow()
    {
        return RowBelow(1);
    }

    public new XlsxTableRow RowBelow(int step)
    {
        return RowShift(step);
    }

    #endregion XlsxTableRow Below

    public new IXlsxTableRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public new IXlsxTableRows InsertRowsAbove(int numberOfRows)
    {
        return XlsxHelper.InsertRowsWithoutEvents(base.InsertRowsAbove, _tableRange, numberOfRows, !_tableRange.Table.ShowTotalsRow);
    }

    public new IXlsxTableRows InsertRowsBelow(int numberOfRows)
    {
        return XlsxHelper.InsertRowsWithoutEvents(base.InsertRowsBelow, _tableRange, numberOfRows, !_tableRange.Table.ShowTotalsRow);
    }

    public new void Delete()
    {
        Delete(XlsxShiftDeletedCells.ShiftCellsUp);
    }
}
