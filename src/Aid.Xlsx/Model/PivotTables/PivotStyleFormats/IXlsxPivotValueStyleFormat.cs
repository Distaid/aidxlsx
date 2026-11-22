// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.PivotTables.PivotValues;

namespace Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

/// <summary>
/// A fluent API for styling a <see cref="IXlsxPivotValue"/> field of a <see cref="IXlsxPivotTable"/>.
/// </summary>
public interface IXlsxPivotValueStyleFormat : IXlsxPivotStyleFormat
{
    /// <summary>
    /// Adds a limitation so the <see cref="IXlsxPivotStyleFormat.Style"/> is only applied to cells in a pivot table
    /// that also belong to the <paramref name="field"/> (label or data).
    /// </summary>
    /// <param name="field">Only cells in a pivot table under this field will be styled.</param>
    IXlsxPivotValueStyleFormat AndWith(IXlsxPivotField field);

    /// <summary>
    /// Adds a limitation so the <see cref="IXlsxPivotStyleFormat.Style"/> is only applied to cells in a pivot table
    /// that also belong to the <paramref name="field"/> data cells. The cell values also must satisfy the <paramref name="predicate"/>.
    /// </summary>
    /// <remarks>
    /// The pivot style is bound by the field index in a pivot table, not field value. E.g. if field values
    /// are Jan, Feb and the predicate marks Feb (offset 1) = second field (Feb) will be highlighted.
    /// If user later reverses order in Excel to Feb, Jan, the style would still apply to the second value - Jan.
    /// </remarks>
    /// <param name="field">Only cells in a pivot table under this field will be styled.</param>
    /// <param name="predicate">A predicate to determine which index of the field should be styled.</param>
    IXlsxPivotValueStyleFormat AndWith(IXlsxPivotField field, Predicate<XlsxCellValue> predicate);

    /// <summary>
    /// Adds a limitation so the <see cref="IXlsxPivotStyleFormat.Style"/> is only applied to cells in a pivot table
    /// that display values for <paramref name="valueField"/> cells (i.e. data cells and grand total).
    /// </summary>
    /// <param name="valueField">One of <see cref="IXlsxPivotTable.Values">value fields</see> of the pivot table.</param>
    IXlsxPivotValueStyleFormat ForValueField(IXlsxPivotValue valueField);
}
