namespace Aid.Xlsx.Model.AutoFilters;

internal class XlsxFilterConnector(XlsxFilterColumn filterColumn) : IXlsxFilterConnector
{
    private readonly XlsxFilterColumn _filterColumn = filterColumn;

    public IXlsxCustomFilteredColumn And => new XlsxCustomFilteredColumn(_filterColumn, XlsxConnector.And);

    public IXlsxCustomFilteredColumn Or => new XlsxCustomFilteredColumn(_filterColumn, XlsxConnector.Or);
}
