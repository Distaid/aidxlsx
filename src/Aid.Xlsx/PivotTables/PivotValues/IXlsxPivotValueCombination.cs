#nullable disable

using Aid.Xlsx.Cells;

namespace Aid.Xlsx.PivotTables.PivotValues
{
    /// <summary>
    /// An interface for fluent configuration of how to show <see cref="IXlsxPivotValue"/>,
    /// when the value should be displayed not as a value itself, but in relation to another
    /// value (e.g. percentage difference in relation to different value).
    /// </summary>
    public interface IXlsxPivotValueCombination
    {
        IXlsxPivotValue And(XlsxCellValue item);

        IXlsxPivotValue AndNext();

        /// <summary>
        /// The base item value for calculation will be the value of the previous row of base
        /// field, depending on the order of base field values in a row/column. If there isn't
        /// a previous value, the same value will be used.
        /// <para>
        /// This only affects display how are values displayed, not the values themselves.
        /// </para>
        /// <para>
        /// Example:
        /// We have a table of sales and a pivot table, where sales are summed per month.
        /// The months are sorted from Jan to Dec. To display a percentage increase of
        /// sales per month (the base value is previous month):
        /// <c>
        /// IXlsxPivotValue sales;
        /// sales.SetSummaryFormula(XlsxPivotSummary.Sum).ShowAsPercentageDifferenceFrom("Month").AndPrevious();
        /// </c>
        /// </para>
        /// </summary>
        IXlsxPivotValue AndPrevious();
    }
}
