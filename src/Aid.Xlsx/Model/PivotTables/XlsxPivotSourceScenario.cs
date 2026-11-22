using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// Source of data for a <see cref="XlsxPivotCache"/> that takes uses scenarios in the workbook to
/// create data.
/// </summary>
internal sealed class XlsxPivotSourceScenario : IXlsxPivotSource
{
    public bool Equals(IXlsxPivotSource? other)
    {
        return other is XlsxPivotSourceScenario;
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxPivotSource other && Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    public bool TryGetSource(XlsxWorkbook workbook, out XlsxWorksheet? sheet, out Area? sheetArea)
    {
        throw new NotImplementedException("Scenario pivot cache data source is not supported.");
    }
}
