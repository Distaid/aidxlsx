using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

/// <summary>
/// A comparator of two cell value. It uses semantic of a sort feature in Excel:
/// <list>
///   <item>Order by type is number, text, logical, error, blank.</item>
///   <item>Errors are not sorted.</item>
///   <item>Blanks are always last, both in ascending and descending order.</item>
///   <item>Stable sort.</item>
/// </list>
/// </summary>
internal class XlsxCellValueSortComparer(IXlsxSortElement sortElement) : IComparer<XlsxCellValue>
{
    private readonly bool _isAscending = sortElement.SortOrder == XlsxSortOrder.Ascending;
    private readonly bool _interpretBlankAsString = !sortElement.IgnoreBlanks;
    private readonly StringComparer _comparer = sortElement.MatchCase ? StringComparer.CurrentCulture : StringComparer.CurrentCultureIgnoreCase;

    public int Compare(XlsxCellValue x, XlsxCellValue y)
    {
        var xTypeOrder = GetTypeOrder(x, _isAscending);
        var yTypeOrder = GetTypeOrder(y, _isAscending);
        if (xTypeOrder != yTypeOrder)
            return xTypeOrder - yTypeOrder;

        return _isAscending ? CompareAsc(x, y) : -CompareAsc(x, y);
    }

    private int GetTypeOrder(XlsxCellValue x, bool asc)
    {
        // Blank is always last, both for asc and desc.
        if (!_interpretBlankAsString && x.Type == XlsxDataType.Blank)
            return 4;

        var ascOrder = x.Type switch
        {
            XlsxDataType.Number => 0,
            XlsxDataType.DateTime => 0,
            XlsxDataType.TimeSpan => 0,
            XlsxDataType.Text => 1,
            XlsxDataType.Blank => 1, // If we get here, the blank is interpreted as a text.
            XlsxDataType.Boolean => 2,
            XlsxDataType.Error => 3,
            _ => throw new NotSupportedException()
        };
        return asc ? ascOrder : -ascOrder;
    }

    private int CompareAsc(XlsxCellValue x, XlsxCellValue y)
    {
        x = _interpretBlankAsString && x.Type == XlsxDataType.Blank ? string.Empty : x;
        y = _interpretBlankAsString && y.Type == XlsxDataType.Blank ? string.Empty : y;
        return x.Type switch
        {
            XlsxDataType.Blank => 0,// Blanks are not sorted. That doesn't really affect content, but cells still contain other info, e.g. style.
            XlsxDataType.Text => _comparer.Compare(x.GetText(), y.GetText()),
            XlsxDataType.Boolean => x.GetBoolean().CompareTo(y.GetBoolean()),
            XlsxDataType.Error => 0,// Errors are never sorted
            XlsxDataType.Number or XlsxDataType.DateTime or XlsxDataType.TimeSpan => x.GetUnifiedNumber().CompareTo(y.GetUnifiedNumber()),
            _ => throw new NotSupportedException(),
        };
    }
}
