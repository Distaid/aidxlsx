using Aid.Xlsx.Cells;

namespace Aid.Xlsx.Ranges.Sort
{
    /// <summary>
    /// A comparator of two cell value. It uses semantic of a sort feature in Excel:
    /// <list>
    ///   <item>Order by type is number, text, logical, error, blank.</item>
    ///   <item>Errors are not sorted.</item>
    ///   <item>Blanks are always last, both in ascending and descending order.</item>
    ///   <item>Stable sort.</item>
    /// </list>
    /// </summary>
    internal class XlsxCellValueSortComparer : IComparer<XlsxCellValue>
    {
        private readonly bool _isAscending;
        private readonly bool _interpretBlankAsString;
        private readonly StringComparer _comparer;

        public XlsxCellValueSortComparer(IXlsxSortElement sortElement)
        {
            // Detecting current culture is expensive, when called enough time. Keep pre-calculated comparer.
            _comparer = sortElement.MatchCase ? StringComparer.CurrentCulture : StringComparer.CurrentCultureIgnoreCase;
            _isAscending = sortElement.SortOrder == XlsxSortOrder.Ascending;
            _interpretBlankAsString = !sortElement.IgnoreBlanks;
        }

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
            switch (x.Type)
            {
                case XlsxDataType.Blank:
                    return 0; // Blanks are not sorted. That doesn't really affect content, but cells still contain other info, e.g. style.

                case XlsxDataType.Text:
                    return _comparer.Compare(x.GetText(), y.GetText());

                case XlsxDataType.Boolean:
                    return x.GetBoolean().CompareTo(y.GetBoolean());

                case XlsxDataType.Error:
                    return 0; // Errors are never sorted

                case XlsxDataType.Number:
                case XlsxDataType.DateTime:
                case XlsxDataType.TimeSpan:
                    return x.GetUnifiedNumber().CompareTo(y.GetUnifiedNumber());

                default:
                    throw new NotSupportedException();
            }
        }
    }
}
