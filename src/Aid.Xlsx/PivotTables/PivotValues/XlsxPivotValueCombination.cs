#nullable disable

using Aid.Xlsx.Cells;

namespace Aid.Xlsx.PivotTables.PivotValues
{
    internal class XlsxPivotValueCombination : IXlsxPivotValueCombination
    {
        private readonly IXlsxPivotValue _pivotValue;

        public XlsxPivotValueCombination(IXlsxPivotValue pivotValue)
        {
            _pivotValue = pivotValue;
        }
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
}
