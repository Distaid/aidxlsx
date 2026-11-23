#nullable disable

using System.Collections;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges;

internal class XlsxRangeColumns(XlsxWorksheet worksheet) : IXlsxRangeColumns
{
    private readonly XlsxWorksheet _worksheet = worksheet;
    private readonly List<XlsxRangeColumn> _ranges = [];

    internal XlsxCellFormat Format
    {
        get
        {
            var columns = _ranges.Select(x => SheetArea.From(x.RangeAddress)).ToArray();
            return XlsxCellFormat.ForAreas(_worksheet.Workbook, columns, null);
        }
    }

    #region IXlsxRangeColumns Members

    public IXlsxStyle Style
    {
        get => Format;
        set => Format.SetStyle(value);
    }

    public IXlsxRangeColumns Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        _ranges.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        _ranges.OrderByDescending(c => c.ColumnNumber()).ForEach(r => r.Delete());
        _ranges.Clear();
    }

    public void Add(IXlsxRangeColumn range)
    {
        _ranges.Add((XlsxRangeColumn)range);
    }

    public IEnumerator<IXlsxRangeColumn> GetEnumerator()
    {
        return _ranges.Cast<IXlsxRangeColumn>()
          .OrderBy(r => r.Worksheet.Position)
          .ThenBy(r => r.ColumnNumber())
          .GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXlsxCells Cells()
    {
        var cells = new XlsxCells(_worksheet, usedCellsOnly: false, options: XlsxCellsUsedOptions.AllContents);
        foreach (XlsxRangeColumn container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public IXlsxCells CellsUsed()
    {
        var cells = new XlsxCells(_worksheet, usedCellsOnly: true, options: XlsxCellsUsedOptions.AllContents);
        foreach (XlsxRangeColumn container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }


    public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
    {
        var cells = new XlsxCells(_worksheet, usedCellsOnly: true, options: options);
        foreach (XlsxRangeColumn container in _ranges)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public void Select()
    {
        foreach (var range in this)
            range.Select();
    }

    #endregion IXlsxRangeColumns Members
}
