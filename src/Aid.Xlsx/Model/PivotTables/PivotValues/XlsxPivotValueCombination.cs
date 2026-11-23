#nullable disable

using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.PivotTables.PivotValues;

internal class XlsxPivotValueCombination(IXlsxPivotValue pivotValue) : IXlsxPivotValueCombination
{
    private readonly IXlsxPivotValue _pivotValue = pivotValue;

    public IXlsxPivotValue And(XlsxCellValue item)
    {
        return _pivotValue
            .SetBaseItemValue(item)
            .SetCalculationItem(XlsxPivotCalculationItem.Value);
    }

    public IXlsxPivotValue AndNext()
    {
        return _pivotValue
            .SetCalculationItem(XlsxPivotCalculationItem.Next);
    }

    public IXlsxPivotValue AndPrevious()
    {
        return _pivotValue
            .SetCalculationItem(XlsxPivotCalculationItem.Previous);
    }
}
