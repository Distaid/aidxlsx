// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.PivotTables.Areas;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.PivotTables.PivotStyleFormats;

/// <summary>
/// A interface for styling various parts of a pivot table, e.g. the whole table, specific
/// area or just a field. Use <see cref="IXlsxPivotTable.StyleFormats"/> and <see cref="IXlsxPivotField.StyleFormats"/>
/// to access it.
/// </summary>
public interface IXlsxPivotStyleFormat
{
    /// <summary>
    /// To what part of the pivot table part will the style apply to.
    /// </summary>
    XlsxPivotStyleFormatElement AppliesTo { get; }

    /// <summary>
    /// The differential style of the part.
    /// </summary>
    /// <remarks>
    /// The final displayed style is done by composing all differential styles that overlap the element.
    /// </remarks>
    IXlsxStyle Style { get; set; }
}
