using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// A class to resolve a format of a cell in a worksheet, if the cell doesn't already has format.
/// It basically goes up the levels and tries to find the "closest" component with a non-empty
/// format. This is a short lived class to be used modifying format of one sheet area.
/// </summary>
internal class FormatResolver(XlsxWorksheet worksheet)
{
    private readonly XlsxCellFormatValue _defaultFormat = worksheet.Workbook.Styles.DefaultFormat;
    private readonly XlsxWorksheet _worksheet = worksheet;
    private readonly XlsxColumnsCollection _columns = worksheet.Internals.ColumnsCollection;
    private readonly XlsxRowsCollection _rows = worksheet.Internals.RowsCollection;

    /// <summary>
    /// Resolve a style of a point in a worksheet according to format hierarchy.
    /// </summary>
    /// <param name="point">Point for which to resolve the style.</param>
    /// <returns>A format that is already registered in the styles.</returns>
    public XlsxCellFormatValue Resolve(Point point)
    {
        if (_rows.TryGetValue(point.Row, out XlsxRow? row) &&
            row.FormatValue is { } rowFormat)
        {
            return rowFormat;
        }

        if (_columns.TryGetValue(point.Column, out XlsxColumn? column) &&
            column.FormatValue is { } columnFormat)
        {
            return columnFormat;
        }

        if (_worksheet.FormatValue is { } worksheetFormat)
            return worksheetFormat;

        return _defaultFormat;
    }
}
