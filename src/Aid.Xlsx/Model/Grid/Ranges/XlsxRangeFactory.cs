using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges;

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
        return key.RangeType switch
        {
            XlsxRangeType.Range => CreateRange(key.RangeAddress),
            XlsxRangeType.Column => CreateColumn(key.RangeAddress.FirstAddress.ColumnNumber),
            XlsxRangeType.Row => CreateColumn(key.RangeAddress.FirstAddress.RowNumber),
            XlsxRangeType.RangeColumn => CreateRangeColumn(key.RangeAddress),
            XlsxRangeType.RangeRow => CreateRangeRow(key.RangeAddress),
            XlsxRangeType.Table => CreateTable(key.RangeAddress),
            _ => throw new NotImplementedException(key.RangeType.ToString()),
        };
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
