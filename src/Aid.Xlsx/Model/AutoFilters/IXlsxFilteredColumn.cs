using Aid.Xlsx.Cells;

namespace Aid.Xlsx.AutoFilters;

/// <summary>
/// A fluent API interface for adding another values to a <see cref="XlsxFilterType.Regular"/>
/// filter. It is chained by <see cref="IXlsxFilterColumn.AddFilter"/> method or
/// <see cref="IXlsxFilterColumn.AddDateGroupFilter"/>.
/// </summary>
public interface IXlsxFilteredColumn
{
    /// <summary>
    /// Add another value to a subset of allowed values for a <see cref="XlsxFilterType.Regular"/>
    /// filter. See <see cref="IXlsxFilterColumn.AddFilter"/> for more details.
    /// </summary>
    /// <param name="value">Value of the filter. The type is <c>XlsxCellValue</c>, but that's for
    /// convenience sake. The value is converted to a string and filter works with string.</param>
    /// <param name="reapply">Should the autofilter be immediately reapplied?</param>
    /// <returns>Fluent API allowing to add additional filter value.</returns>
    IXlsxFilteredColumn AddFilter(XlsxCellValue value, bool reapply = true);

    /// <summary>
    /// Add another grouping to a set of allowed groupings. See <see cref="IXlsxFilterColumn.AddDateGroupFilter"/>
    /// for more details.
    /// </summary>
    /// <returns>Fluent API allowing to add additional date group filter.</returns>
    IXlsxFilteredColumn AddDateGroupFilter(DateTime date, XlsxDateTimeGrouping dateTimeGrouping, bool reapply = true);
}
