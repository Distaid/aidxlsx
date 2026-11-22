namespace Aid.Xlsx.AutoFilters
{
    public interface IXlsxFilterConnector
    {
        IXlsxCustomFilteredColumn And { get; }
        IXlsxCustomFilteredColumn Or { get; }
    }
}
