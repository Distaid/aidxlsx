using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

internal class XlsxRangeColumnsSortComparer : IComparer<int>
{
    private readonly List<(int RowNumber, XlsxCellValueSortComparer Comparer)> _rowComparers;
    private readonly ValueSlice _valueSlice;

    internal XlsxRangeColumnsSortComparer(XlsxWorksheet sheet, Area sortRange, IXlsxSortElements sortRows)
    {
        if (!sortRows.Any())
            throw new ArgumentException("Empty sort specification.");

        if (sortRange.Width < sortRows.Max(x => x.ElementNumber))
            throw new ArgumentException("Range has fewer columns that sort specification.");

        _valueSlice = sheet.Internals.CellsCollection.ValueSlice;
        _rowComparers = [.. sortRows.Select(se => (se.ElementNumber + sortRange.TopRow - 1, new XlsxCellValueSortComparer(se)))];
    }

    public int Compare(int colNumber1, int colNumber2)
    {
        foreach (var (rowNumber, comparer) in _rowComparers)
        {
            var col1 = _valueSlice.GetCellValue(new Point(rowNumber, colNumber1));
            var col2 = _valueSlice.GetCellValue(new Point(rowNumber, colNumber2));
            var comparison = comparer.Compare(col1, col2);
            if (comparison != 0)
                return comparison;
        }

        // Workaround for stable sort, see XlsxRangeRowsSortComparer.
        return colNumber1 - colNumber2;
    }
}
