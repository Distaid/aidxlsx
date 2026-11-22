namespace Aid.Xlsx.AutoFilters;

internal class XlsxFilterConnector : IXlsxFilterConnector
{
    private readonly XlsxFilterColumn _filterColumn;

    public XlsxFilterConnector(XlsxFilterColumn filterColumn)
    {
        _filterColumn = filterColumn;
    }

    public IXlsxCustomFilteredColumn And => new XlsxCustomFilteredColumn(_filterColumn, XlsxConnector.And);

    public IXlsxCustomFilteredColumn Or => new XlsxCustomFilteredColumn(_filterColumn, XlsxConnector.Or);
}
