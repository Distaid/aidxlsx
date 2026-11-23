#nullable disable

namespace Aid.Xlsx.Model.Drawings;

internal class XlsxDrawingPosition : IXlsxDrawingPosition
{
    public int Column { get; set; }
    public IXlsxDrawingPosition SetColumn(int column) { Column = column; return this; }
    public double ColumnOffset { get; set; }
    public IXlsxDrawingPosition SetColumnOffset(double columnOffset) { ColumnOffset = columnOffset; return this; }

    public int Row { get; set; }
    public IXlsxDrawingPosition SetRow(int row) { Row = row; return this; }
    public double RowOffset { get; set; }
    public IXlsxDrawingPosition SetRowOffset(double rowOffset) { RowOffset = rowOffset; return this; }
}
