using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.PivotTables;

/// <summary>
/// An abstraction of source data for a <see cref="XlsxPivotCache"/>. Implementations must correctly
/// implement equals.
/// </summary>
internal interface IXlsxPivotSource : IEquatable<IXlsxPivotSource>
{
    /// <summary>
    /// Try to determine actual area of the source reference in the
    /// workbook. Source reference might not be valid in the workbook, some might
    /// not be supported.
    /// </summary>
    bool TryGetSource(XlsxWorkbook workbook, out XlsxWorksheet? sheet, out Area? sheetArea);
}
