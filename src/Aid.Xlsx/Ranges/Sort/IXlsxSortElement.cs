using Aid.Xlsx.Cells;

namespace Aid.Xlsx.Ranges.Sort
{
    public enum XlsxSortOrder { Ascending, Descending }
    public enum XlsxSortOrientation { TopToBottom, LeftToRight }
    public interface IXlsxSortElement
    {
        /// <summary>
        /// Column or row number whose values will be used for sorting.
        /// </summary>
        Int32 ElementNumber { get; }

        /// <summary>
        /// Sorting order.
        /// </summary>
        XlsxSortOrder SortOrder { get; }

        /// <summary>
        /// When <c>true</c> (recommended, matches Excel behavior), blank cell values are always
        /// sorted at the end regardless of sorting order. When <c>false</c>, blank values are
        /// considered empty strings and are sorted among other cell values with a type
        /// <see cref="XlsxDataType.Text"/>.
        /// </summary>
        Boolean IgnoreBlanks { get; }

        /// <summary>
        /// When cell value is a <see cref="XlsxDataType.Text"/>, should sorting be case insensitive
        /// (<c>false</c>, Excel default behavior) or case sensitive (<c>true</c>). Doesn't affect
        /// other cell value types.
        /// </summary>
        Boolean MatchCase { get; }
    }
}
