using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Calculation.Exceptions;

/// <summary>
/// Exception that happens when formula in a cell depends on other cells,
/// but the supporting formulas are still dirty.
/// </summary>
internal class GettingDataException(SheetPoint point) : Exception
{
    public SheetPoint Point { get; } = point;
}
