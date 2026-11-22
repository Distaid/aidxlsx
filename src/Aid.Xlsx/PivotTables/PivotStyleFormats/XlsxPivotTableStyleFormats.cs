// Keep this file CodeMaid organised and cleaned
namespace Aid.Xlsx.PivotTables.PivotStyleFormats;

internal class XlsxPivotTableStyleFormats : IXlsxPivotTableStyleFormats
{
    private readonly XlsxPivotTable _pivotTable;

    public XlsxPivotTableStyleFormats(XlsxPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    #region IXlsxPivotTableStyleFormats members

    public IXlsxPivotStyleFormats ColumnGrandTotalFormats => new XlsxPivotStyleFormats(_pivotTable, isRowGrand: false);

    public IXlsxPivotStyleFormats RowGrandTotalFormats => new XlsxPivotStyleFormats(_pivotTable, isRowGrand: true);

    #endregion IXlsxPivotTableStyleFormats members
}
