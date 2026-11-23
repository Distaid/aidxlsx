using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.PivotTables.Areas;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// A field displayed in the filters part of a pivot table.
/// </summary>
internal class XlsxPivotPageField
{
    internal XlsxPivotPageField(int field)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(field);

        Field = field;
    }

    /// <summary>
    /// Field index to <see cref="XlsxPivotTable.PivotFields"/>. Can't contain
    /// <see cref="XlsxConstants.PivotTable.ValuesSentinalLabel">'data'</see>
    /// field <c>-2</c>.
    /// </summary>
    internal int Field { get; }

    /// <summary>
    /// If a single item is selected, item index. Null, if nothing selected or multiple selected.
    /// Multiple selected values are indicated directly in <see cref="XlsxPivotTableField.Items"/>
    /// through <see cref="XlsxPivotFieldItem.Hidden"/> flags. Items that are not selected are hidden,
    /// rest isn't.
    /// </summary>
    internal int? ItemIndex { get; set; }

    // OLAP
    internal int? HierarchyIndex { get; init; }

    // OLAP
    internal string? HierarchyUniqueName { get; init; }

    // OLAP
    internal string? HierarchyDisplayName { get; init; }
}
