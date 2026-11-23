#nullable disable

namespace Aid.Xlsx.Model.Drawings;

public interface IXlsxDrawingPosition
{
    int Column { get; set; }
    IXlsxDrawingPosition SetColumn(int column);
    double ColumnOffset { get; set; }
    IXlsxDrawingPosition SetColumnOffset(double columnOffset);

    int Row { get; set; }
    IXlsxDrawingPosition SetRow(int row);
    double RowOffset { get; set; }
    IXlsxDrawingPosition SetRowOffset(double rowOffset);
}
