namespace Aid.Xlsx.Model.AutoFilters;

public interface IXlsxFilterConnector
{
    IXlsxCustomFilteredColumn And { get; }
    IXlsxCustomFilteredColumn Or { get; }
}
