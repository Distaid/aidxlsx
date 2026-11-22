#nullable disable

namespace Aid.Xlsx.Cells
{
    internal class XlsxCellValueComparer : IEqualityComparer<XlsxCellValue>
    {
        private readonly StringComparer _textComparer;

        internal static readonly XlsxCellValueComparer OrdinalIgnoreCase = new(StringComparer.OrdinalIgnoreCase);

        private XlsxCellValueComparer(StringComparer textComparer)
        {
            _textComparer = textComparer;
        }

        public bool Equals(XlsxCellValue x, XlsxCellValue y)
        {
            if (x.Type != y.Type)
                return false;

            return x.Type switch
            {
                XlsxDataType.Blank => true,
                XlsxDataType.Boolean => x.GetBoolean() == y.GetBoolean(),
                XlsxDataType.Number => x.GetNumber().Equals(y.GetNumber()),
                XlsxDataType.Text => _textComparer.Equals(x.GetText(), y.GetText()),
                XlsxDataType.Error => x.GetError() == y.GetError(),
                XlsxDataType.DateTime => x.GetUnifiedNumber().Equals(y.GetUnifiedNumber()),
                XlsxDataType.TimeSpan => x.GetUnifiedNumber().Equals(y.GetUnifiedNumber()),
                _ => throw new NotSupportedException()
            };
        }

        public int GetHashCode(XlsxCellValue obj)
        {
            unchecked
            {
                var hashCode = obj.Type.GetHashCode();
                var valueHashCode = obj.Type switch
                {
                    XlsxDataType.Blank => 0,
                    XlsxDataType.Boolean => obj.GetBoolean().GetHashCode(),
                    XlsxDataType.Text => _textComparer.GetHashCode(obj.GetText()),
                    XlsxDataType.Error => obj.GetError().GetHashCode(),
                    _ => obj.GetUnifiedNumber().GetHashCode()
                };
                hashCode = (hashCode * 397) ^ valueHashCode;
                return hashCode;
            }
        }
    }
}
