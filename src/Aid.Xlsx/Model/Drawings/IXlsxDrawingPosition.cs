#nullable disable

namespace Aid.Xlsx.Drawings
{
    public interface IXlsxDrawingPosition
    {
        Int32 Column { get; set; }
        IXlsxDrawingPosition SetColumn(Int32 column);
        Double ColumnOffset { get; set; }
        IXlsxDrawingPosition SetColumnOffset(Double columnOffset);

        Int32 Row { get; set; }
        IXlsxDrawingPosition SetRow(Int32 row);
        Double RowOffset { get; set; }
        IXlsxDrawingPosition SetRowOffset(Double rowOffset);
    }
}
