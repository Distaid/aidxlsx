using Aid.Xlsx.Columns;
using Aid.Xlsx.Rows;
using Aid.Xlsx.Tables;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Ranges
{
    internal class XlsxRangeFactory
    {
        private readonly XlsxWorksheet _worksheet;

        public XlsxRangeFactory(XlsxWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        #region Methods

        public XlsxRangeBase Create(XlsxRangeKey key)
        {
            switch (key.RangeType)
            {
                case XlsxRangeType.Range:
                    return CreateRange(key.RangeAddress);

                case XlsxRangeType.Column:
                    return CreateColumn(key.RangeAddress.FirstAddress.ColumnNumber);

                case XlsxRangeType.Row:
                    return CreateColumn(key.RangeAddress.FirstAddress.RowNumber);

                case XlsxRangeType.RangeColumn:
                    return CreateRangeColumn(key.RangeAddress);

                case XlsxRangeType.RangeRow:
                    return CreateRangeRow(key.RangeAddress);

                case XlsxRangeType.Table:
                    return CreateTable(key.RangeAddress);

                case XlsxRangeType.Worksheet:
                default:
                    throw new NotImplementedException(key.RangeType.ToString());
            }
        }

        public XlsxRange CreateRange(XlsxRangeAddress rangeAddress)
        {
            return new XlsxRange(rangeAddress, _worksheet.Style);
        }

        public XlsxColumn CreateColumn(int columnNumber)
        {
            return new XlsxColumn(_worksheet, columnNumber);
        }

        public XlsxRow CreateRow(int rowNumber)
        {
            return new XlsxRow(_worksheet, rowNumber);
        }

        public XlsxRangeColumn CreateRangeColumn(XlsxRangeAddress rangeAddress)
        {
            return new XlsxRangeColumn(rangeAddress);
        }

        public XlsxRangeRow CreateRangeRow(XlsxRangeAddress rangeAddress)
        {
            return new XlsxRangeRow(rangeAddress);
        }

        public XlsxTable CreateTable(XlsxRangeAddress rangeAddress)
        {
            return new XlsxTable(rangeAddress, _worksheet.Style);
        }

        #endregion Methods
    }
}
