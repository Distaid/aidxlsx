#nullable disable

using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Calculation;

internal class XlsxAddressComparer : IEqualityComparer<IXlsxAddress>
{
    private readonly bool _ignoreFixed;

    public XlsxAddressComparer(bool ignoreFixed)
    {
        _ignoreFixed = ignoreFixed;
    }

    public bool Equals(IXlsxAddress x, IXlsxAddress y)
    {
        return (x == null && y == null) ||
               (x != null && y != null &&
                string.Equals(x.Worksheet.Name, y.Worksheet.Name, StringComparison.InvariantCultureIgnoreCase) &&
                x.ColumnNumber == y.ColumnNumber &&
                x.RowNumber == y.RowNumber &&
                (_ignoreFixed || x.FixedColumn == y.FixedColumn &&
                 x.FixedRow == y.FixedRow));
    }

    public int GetHashCode(IXlsxAddress obj)
    {
        return new
        {
            WorksheetName = obj.Worksheet.Name.ToUpperInvariant(),
            obj.ColumnNumber,
            obj.RowNumber,
            FixedColumn = (!_ignoreFixed && obj.FixedColumn),
            FixedRow = (!_ignoreFixed && obj.FixedRow)
        }.GetHashCode();
    }
}
