#nullable disable

using System.Collections;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges;

internal class XlsxRangeRows(XlsxWorksheet worksheet) : IXlsxRangeRows
{
    private readonly XlsxWorksheet _worksheet = worksheet;
    private readonly List<XlsxRangeRow> _ranges = [];

    internal XlsxCellFormat Format
    {
        get
        {
            SheetArea[] areas = [.. _ranges.Select(x => SheetArea.From(x.RangeAddress))];
            return XlsxCellFormat.ForAreas(_worksheet.Workbook, areas, null);
        }
    }

    #region IXlsxRangeRows Members

    public IXlsxStyle Style
    {
        get => Format;
        set => Format.SetStyle(value);
    }

    public IXlsxRangeRows Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        _ranges.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
        _ranges.Clear();
    }

    public void Add(IXlsxRangeRow range)
    {
        _ranges.Add((XlsxRangeRow)range);
    }

    public IEnumerator<IXlsxRangeRow> GetEnumerator()
    {
        return _ranges.Cast<IXlsxRangeRow>()
                      .OrderBy(r => r.Worksheet.Position)
                      .ThenBy(r => r.RowNumber())
                      .GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXlsxCells Cells()
    {
        var cells = new XlsxCells(_worksheet, false, XlsxCellsUsedOptions.AllContents);
        foreach (XlsxRangeRow container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public IXlsxCells CellsUsed()
    {
        var cells = new XlsxCells(_worksheet, true, XlsxCellsUsedOptions.AllContents);
        foreach (XlsxRangeRow container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }


    public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
    {
        var cells = new XlsxCells(_worksheet, true, options);
        foreach (XlsxRangeRow container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public void Select()
    {
        foreach (IXlsxRangeRow range in this)
            range.Select();
    }

    #endregion IXlsxRangeRows Members
}
