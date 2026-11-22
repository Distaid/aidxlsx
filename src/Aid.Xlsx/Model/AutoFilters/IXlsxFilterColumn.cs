using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Rows;

namespace Aid.Xlsx.Model.AutoFilters;

/// <summary>
/// Type of a <see cref="XlsxFilterType.TopBottom"/> filter that is used to determine number of
/// visible top/bottom values.
/// </summary>
public enum XlsxTopBottomType
{
    /// <summary>
    /// Filter should display requested number of items.
    /// </summary>
    Items,

    /// <summary>
    /// Number of displayed items is determined as a percentage data rows of auto filter.
    /// </summary>
    Percent
}

public enum XlsxDateTimeGrouping { Year, Month, Day, Hour, Minute, Second }

/// <summary>
/// <para>
/// AutoFilter filter configuration for one column in an autofilter <see cref="IXlsxAutoFilter.Range">area</see>.
/// Filters determine visibility of rows in the autofilter area. Value in the row must satisfy
/// all filters in all columns in order for row to be visible, otherwise it is <see cref="IXlsxRow.IsHidden"/>.
/// </para>
/// <para>
/// Column can have only one type of filter, so it's not possible to combine several different
/// filter types on one column. Methods for adding filters clear other types or remove
/// previously set filters when needed. Some types of filters can have multiple conditions (e.g.
/// <see cref="XlsxFilterType.Regular"/> can have many values while <see cref="XlsxFilterType.Dynamic"/>
/// can be only one).
/// </para>
/// <para>
/// Whenever filter configuration changes, the filters are immediately reapplied.
/// </para>
/// </summary>
public interface IXlsxFilterColumn
{
    /// <summary>
    /// Remove all filters from the column.
    /// </summary>
    /// <param name="reapply">Should the autofilter be immediately reapplied?</param>
    void Clear(bool reapply = true);

    /// <summary>
    /// <para>
    /// Switch to the <see cref="XlsxFilterType.Regular"/> filter if filter column has a
    /// different type (for current type <see cref="FilterType"/>) and add
    /// <paramref name="value"/> to a set of allowed values. Excel displays regular filter as
    /// a list of possible values in a column with checkbox next to it and user can check which
    /// one should be displayed.
    /// </para>
    /// <para>
    /// From technical perspective, the passed <paramref name="value"/> is converted to
    /// a localized string (using current locale) and the column values satisfy the filter
    /// condition, when the <see cref="IXlsxCell.GetFormattedString">formatted string of a cell
    /// </see> matches any filter string.
    /// </para>
    /// <para>
    /// Examples of less intuitive behavior: filter value is <c>2.5</c> in locale cs-CZ that
    /// uses "<em>,</em>" as a decimal separator. The passed <paramref name="value"/>
    /// is number 2.5, converted immediately to a string <em>2,5</em>. The string is used for
    /// comparison with values of cells in the column:
    /// <list type="bullet">
    ///  <item>Number 2.5 formatted with two decimal places as <em>2,50</em> will not match.</item>
    ///  <item>Number 2.5 with default formatting will be matched, because its string is
    ///        <em>2,5</em> in cs-CZ locale (but not in others, e.g. en-US locale).</item>
    ///  <item>Text <em>2,5</em> will be matched.</item>
    /// </list>
    /// </para>
    /// </summary>
    /// <remarks>
    /// This behavior of course highly depends on locale and working with same file on two
    /// different locales might lead to different results.
    /// </remarks>
    /// <param name="value">Value of the filter. The type is <c>XlsxCellValue</c>, but that's for
    /// convenience sake. The value is converted to a string and filter works with string.</param>
    /// <param name="reapply">Should the autofilter be immediately reapplied?</param>
    /// <returns>Fluent API allowing to add additional filter value.</returns>
    IXlsxFilteredColumn AddFilter(XlsxCellValue value, bool reapply = true);

    /// <summary>
    /// <para>
    /// Enable autofilter (if needed), switch to the <see cref="XlsxFilterType.Regular"/> filter
    /// if filter column has a different type (for current type <see cref="FilterType"/>) and
    /// add a filter that is satisfied when cell value is a <see cref="XlsxDataType.DateTime"/>
    /// and the tested date has same components from <paramref name="dateTimeGrouping"/>
    /// component up to the <see cref="XlsxDateTimeGrouping.Year"/> component with same value
    /// as the <paramref name="dateTimeGrouping"/>.
    /// </para>
    /// <para>
    /// The condition basically defines a date range (based on the <paramref name="dateTimeGrouping"/>)
    /// and all dates in the range satisfy the filter. If condition is a day, all date-times
    /// in the day satisfy the filter. If condition is a month, all date-times in the month
    /// satisfy the filter.
    /// </para>
    /// <para>
    /// <example>
    /// Example:
    /// <code>
    /// // Filter will be satisfied if the cell value is a XlsxDataType.DateTime and the month,
    /// // and year are same as the passed date. The day component in the <c>DateTime</c>
    /// // is ignored
    /// AddDateGroupFilter(new DateTime(2023, 7, 15), XlsxDateTimeGrouping.Month)
    /// </code>
    /// </example>
    /// </para>
    /// <para>
    /// There can be multiple date group filters and they are <see cref="XlsxFilterType.Regular"/>
    /// filter types, i.e. they don't delete filters from <see cref="AddFilter"/>. The cell
    /// value is satisfied, if it matches any of the text values from <see cref="AddFilter"/>
    /// or any date group filter.
    /// </para>
    /// </summary>
    /// <param name="date">Date which components are compared with date values of the column.</param>
    /// <param name="dateTimeGrouping">
    /// Starting component of the grouping. Tested date must match all date components of the
    /// <paramref name="date"/> from this one to the <see cref="XlsxDateTimeGrouping.Year"/>.
    /// </param>
    /// <param name="reapply">Should the autofilter be immediately reapplied?</param>
    /// <returns>Fluent API allowing to add additional date time group value.</returns>
    IXlsxFilteredColumn AddDateGroupFilter(DateTime date, XlsxDateTimeGrouping dateTimeGrouping, bool reapply = true);

    /// <exception cref="ArgumentOutOfRangeException">If <paramref name="value"/> is out of range 1..500.</exception>
    void Top(Int32 value, XlsxTopBottomType type = XlsxTopBottomType.Items, bool reapply = true);

    /// <exception cref="ArgumentOutOfRangeException">If <paramref name="value"/> is out of range 1..500.</exception>
    void Bottom(Int32 value, XlsxTopBottomType type = XlsxTopBottomType.Items, bool reapply = true);

    void AboveAverage(bool reapply = true);

    void BelowAverage(bool reapply = true);

    IXlsxFilterConnector EqualTo(XlsxCellValue value, bool reapply = true);

    IXlsxFilterConnector NotEqualTo(XlsxCellValue value, bool reapply = true);

    IXlsxFilterConnector GreaterThan(XlsxCellValue value, bool reapply = true);

    IXlsxFilterConnector LessThan(XlsxCellValue value, bool reapply = true);

    IXlsxFilterConnector EqualOrGreaterThan(XlsxCellValue value, bool reapply = true);

    IXlsxFilterConnector EqualOrLessThan(XlsxCellValue value, bool reapply = true);

    void Between(XlsxCellValue minValue, XlsxCellValue maxValue, bool reapply = true);

    void NotBetween(XlsxCellValue minValue, XlsxCellValue maxValue, bool reapply = true);

    IXlsxFilterConnector BeginsWith(String value, bool reapply = true);

    IXlsxFilterConnector NotBeginsWith(String value, bool reapply = true);

    IXlsxFilterConnector EndsWith(String value, bool reapply = true);

    IXlsxFilterConnector NotEndsWith(String value, bool reapply = true);

    IXlsxFilterConnector Contains(String value, bool reapply = true);

    IXlsxFilterConnector NotContains(String value, bool reapply = true);

    /// <summary>
    /// Current filter type used by the filter columns.
    /// </summary>
    XlsxFilterType FilterType { get; }

    /// <summary>
    /// Configuration of a <see cref="XlsxFilterType.TopBottom"/> filter. It contains how many
    /// items/percent (depends on <see cref="TopBottomType"/>) should filter accept.
    /// </summary>
    /// <remarks>
    /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XlsxFilterType.TopBottom"/>.
    /// </remarks>
    Int32 TopBottomValue { get; }

    /// <summary>
    /// Configuration of a <see cref="XlsxFilterType.TopBottom"/> filter. It contains the content
    /// interpretation of a <see cref="TopBottomValue"/> property, i.e. does it mean how many
    /// percents or how many items?
    /// </summary>
    /// <remarks>
    /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XlsxFilterType.TopBottom"/>.
    /// </remarks>
    XlsxTopBottomType TopBottomType { get; }

    /// <summary>
    /// Configuration of a <see cref="XlsxFilterType.TopBottom"/> filter. It determines if filter
    /// should accept items from top or bottom.
    /// </summary>
    /// <remarks>
    /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XlsxFilterType.TopBottom"/>.
    /// </remarks>
    XlsxTopBottomPart TopBottomPart { get; }

    /// <summary>
    /// Configuration of a <see cref="XlsxFilterType.Dynamic"/> filter. It determines the type of
    /// dynamic filter.
    /// </summary>
    /// <remarks>
    /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XlsxFilterType.Dynamic"/>.
    /// </remarks>
    XlsxFilterDynamicType DynamicType { get; }

    /// <summary>
    /// Configuration of a <see cref="XlsxFilterType.Dynamic"/> filter. It contains the dynamic
    /// value used by the filter, e.g. average. The interpretation depends on
    /// <see cref="DynamicType"/>.
    /// </summary>
    /// <remarks>
    /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XlsxFilterType.Dynamic"/>.
    /// </remarks>
    Double DynamicValue { get; }
}
