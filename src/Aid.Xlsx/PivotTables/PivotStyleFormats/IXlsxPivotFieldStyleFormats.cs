// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.PivotTables.Areas;

namespace Aid.Xlsx.PivotTables.PivotStyleFormats;

/// <summary>
/// Interface to change the style of a <see cref="IXlsxPivotField"/> or its parts.
/// </summary>
public interface IXlsxPivotFieldStyleFormats
{
    /// <summary>
    /// Pivot table style of the field values displayed in the data area of the pivot table.
    /// </summary>
    IXlsxPivotValueStyleFormat DataValuesFormat { get; }

    /// <summary>
    /// Get the style of the pivot field header. The head usually contains a name of the field.
    /// In some layouts, header is not individually displayed (e.g. compact), while in others
    /// it is (e.g. tabular).
    /// </summary>
    IXlsxPivotStyleFormat Header { get; }

    /// <summary>
    /// Get the style of the pivot field label values on horizontal or vertical axis.
    /// </summary>
    IXlsxPivotStyleFormat Label { get; }

    IXlsxPivotStyleFormat Subtotal { get; }
}
