using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.PivotTables;

/// <summary>
/// Source of data for a <see cref="XlsxPivotCache"/> that takes data from a connection
/// to external source of data (e.g. database or a workbook).
/// </summary>
internal sealed class XlsxPivotSourceConnection : IXlsxPivotSource
{
    public XlsxPivotSourceConnection(uint connectionId)
    {
        ConnectionId = connectionId;
    }

    public uint ConnectionId { get; }

    public bool Equals(IXlsxPivotSource? otherSource)
    {
        var other = otherSource as XlsxPivotSourceConnection;
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return ConnectionId == other.ConnectionId;
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxPivotSource other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(ConnectionId).GetHashCode();
    }

    public bool TryGetSource(XlsxWorkbook workbook, out XlsxWorksheet? sheet, out Area? sheetArea)
    {
        throw new NotImplementedException("Pivot cache source using a connection is not supported.");
    }
}
