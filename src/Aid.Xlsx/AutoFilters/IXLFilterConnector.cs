namespace Aid.Xlsx
{
    public interface IXLFilterConnector
    {
        IXLCustomFilteredColumn And { get; }
        IXLCustomFilteredColumn Or { get; }
    }
}
