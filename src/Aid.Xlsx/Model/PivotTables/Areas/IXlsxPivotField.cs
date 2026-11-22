#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.PivotTables.PivotStyleFormats;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.PivotTables.Areas
{
    public enum XlsxSubtotalFunction
    {
        Automatic,
        Sum,
        Count,
        Average,
        Maximum,
        Minimum,
        Product,
        CountNumbers,
        StandardDeviation,
        PopulationStandardDeviation,
        Variance,
        PopulationVariance,
    }

    public enum XlsxPivotLayout { Outline, Tabular, Compact }

    /// <summary>
    /// A fluent API representation of a field on an <see cref="IXlsxPivotTable.RowLabels">row</see>,
    /// <see cref="IXlsxPivotTable.ColumnLabels">column</see> or <see cref="IXlsxPivotTable.ReportFilters">
    /// filter</see> axis of a <see cref="IXlsxPivotTable"/>.
    /// </summary>
    /// <remarks>
    /// If the field is a 'data' field, a lot of properties don't make sense and can't be set. In
    /// such case, the setter will throw <exception cref="InvalidOperationException"/> and getter
    /// will return default value for the field.
    /// </remarks>
    public interface IXlsxPivotField
    {
        /// <summary>
        /// <para>
        /// Name of the field in a pivot table <see cref="IXlsxPivotTable.PivotCache"/>. If the field
        /// is 'data' field, return <see cref="XlsxConstants.PivotTable.ValuesSentinalLabel"/>.
        /// </para>
        /// <para>
        /// Note that field name in pivot cache is generally same as in the source data range, but
        /// not always. Field names are unique in the cache and if the source data range contains
        /// duplicate column names, the cache will rename them to keep all names unique.
        /// </para>
        /// </summary>
        String SourceName { get; }

        /// <summary>
        /// <see cref="CustomName"/> of the field in the pivot table. Custom name is a unique
        /// across all fields used in the pivot table (e.g. if same field is added to values area
        /// multiple times, it must have custom name, e.g. <c>Sum1 of Field</c>,
        /// <c>Sum2 of Field</c>).
        /// </summary>
        /// <exception cref="ArgumentException">When setting name to a name that is already used by
        ///     another field.</exception>
        String CustomName { get; set; }

        String SubtotalCaption { get; set; }

        /// <summary>
        /// Get subtotals of the field. The content of the collection depends on the type of subtotal:
        /// <list type="bullet">
        ///   <item>None - the collection is empty.</item>
        ///   <item>Automatic - the collection contains one element with function <see cref="XlsxSubtotalFunction.Automatic"/>.</item>
        ///   <item>Custom - the collection contains a set of functions (at least one) except the <see cref="XlsxSubtotalFunction.Automatic"/>.</item>
        /// </list>
        /// </summary>
        IReadOnlyCollection<XlsxSubtotalFunction> Subtotals { get; }
        Boolean IncludeNewItemsInFilter { get; set; }
        Boolean Outline { get; set; }
        Boolean Compact { get; set; }
        Boolean? SubtotalsAtTop { get; set; }
        Boolean RepeatItemLabels { get; set; }
        Boolean InsertBlankLines { get; set; }
        Boolean ShowBlankItems { get; set; }
        Boolean InsertPageBreaks { get; set; }

        /// <summary>
        /// Are all items of the field collapsed?
        /// </summary>
        /// <remarks>
        /// If only a subset of items is collapsed, getter returns <c>false</c>.
        /// </remarks>
        Boolean Collapsed { get; set; }

        XlsxPivotSortType SortType { get; set; }

        /// <inheritdoc cref="CustomName"/>
        IXlsxPivotField SetCustomName(String value);

        IXlsxPivotField SetSubtotalCaption(String value);

        IXlsxPivotField AddSubtotal(XlsxSubtotalFunction value);

        IXlsxPivotField SetIncludeNewItemsInFilter(Boolean value = true);

        IXlsxPivotField SetLayout(XlsxPivotLayout value);

        IXlsxPivotField SetSubtotalsAtTop(Boolean value = true);

        IXlsxPivotField SetRepeatItemLabels(Boolean value = true);

        IXlsxPivotField SetInsertBlankLines(Boolean value = true);

        IXlsxPivotField SetShowBlankItems(Boolean value = true);

        IXlsxPivotField SetInsertPageBreaks(Boolean value = true);

        IXlsxPivotField SetCollapsed(Boolean value = true);

        IXlsxPivotField SetSort(XlsxPivotSortType value);

        /// <summary>
        /// Selected values for <see cref="IXlsxPivotTable.ReportFilters"/> filter of the pivot
        /// table. Empty for non-filter fields.
        /// </summary>
        IReadOnlyList<XlsxCellValue> SelectedValues { get; }

        /// <summary>
        /// Add a value to selected values of a filter field (<see cref="IXlsxPivotTable.ReportFilters"/>).
        /// Doesn't do anything, if this field is not a filter fields.
        /// </summary>
        IXlsxPivotField AddSelectedValue(XlsxCellValue value);

        /// <summary>
        /// Add a values to a selected values of a filter field. Doesn't do anything if this field
        /// is not a filter fields.
        /// </summary>
        IXlsxPivotField AddSelectedValues(IEnumerable<XlsxCellValue> values);

        IXlsxPivotFieldStyleFormats StyleFormats { get; }

        Boolean IsOnRowAxis { get; }
        Boolean IsOnColumnAxis { get; }
        Boolean IsInFilterList { get; }

        /// <summary>
        /// Index of a field in <see cref="XlsxPivotTable.PivotFields">all pivot fields</see> or -2
        /// for data field.
        /// </summary>
        Int32 Offset { get; }
    }
}
