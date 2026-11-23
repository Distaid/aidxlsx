using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// Source of data for a <see cref="XlsxPivotCache"/> that takes data from a connection
/// to external source of data (e.g. database or a workbook).
/// </summary>
internal sealed class XlsxPivotSourceConnection(uint connectionId) : IXlsxPivotSource
{
    public uint ConnectionId { get; } = connectionId;

    public bool Equals(IXlsxPivotSource? otherSource)
    {
        if (otherSource is not XlsxPivotSourceConnection other)
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
