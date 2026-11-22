#nullable disable

namespace Aid.Xlsx.Drawings
{
    internal class XlsxDrawingPosition : IXlsxDrawingPosition
    {
        public Int32 Column { get; set; }
        public IXlsxDrawingPosition SetColumn(Int32 column) { Column = column; return this; }
        public Double ColumnOffset { get; set; }
        public IXlsxDrawingPosition SetColumnOffset(Double columnOffset) { ColumnOffset = columnOffset; return this; }

        public Int32 Row { get; set; }
        public IXlsxDrawingPosition SetRow(Int32 row) { Row = row; return this; }
        public Double RowOffset { get; set; }
        public IXlsxDrawingPosition SetRowOffset(Double rowOffset) { RowOffset = rowOffset; return this; }
    }
}
