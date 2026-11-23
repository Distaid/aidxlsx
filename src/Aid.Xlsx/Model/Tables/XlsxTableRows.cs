using System.Collections;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Tables;

internal class XlsxTableRows : IXlsxTableRows
{
    private readonly XlsxWorksheet _worksheet;
    private readonly List<XlsxTableRow> _ranges = [];

    public XlsxTableRows(XlsxWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    #region IXlsxTableRows Members

    public IXlsxStyle Style
    {
        get => Format;
        set => Format.SetStyle(value);
    }

    public IXlsxTableRows Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        _ranges.ForEach(r => r.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
        _ranges.Clear();
    }

    public void Add(IXlsxTableRow tableRow)
    {
        _ranges.Add((XlsxTableRow)tableRow);
    }

    public IEnumerator<IXlsxTableRow> GetEnumerator()
    {
        var retList = new List<IXlsxTableRow>();
        _ranges.ForEach(retList.Add);
        return retList.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXlsxCells Cells()
    {
        var cells = new XlsxCells(_worksheet, false, XlsxCellsUsedOptions.AllContents);
        foreach (XlsxTableRow container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public IXlsxCells CellsUsed()
    {
        var cells = new XlsxCells(_worksheet, true, XlsxCellsUsedOptions.AllContents);
        foreach (XlsxTableRow container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public IXlsxCells CellsUsed(Boolean includeFormats)
    {
        return CellsUsed(includeFormats
            ? XlsxCellsUsedOptions.All
            : XlsxCellsUsedOptions.AllContents);
    }

    public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
    {
        var cells = new XlsxCells(_worksheet, false, options);
        foreach (XlsxTableRow container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public void Select()
    {
        foreach (var range in this)
            range.Select();
    }

    #endregion IXlsxTableRows Members

    internal XlsxCellFormat Format
    {
        get
        {
            var rowAreas = _ranges.Select(x => SheetArea.From(x.RangeAddress)).ToArray();
            return XlsxCellFormat.ForTableRows(_worksheet, rowAreas);
        }
    }
}
