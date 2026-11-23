// Keep this file CodeMaid organised and cleaned
namespace Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

internal class XlsxPivotTableStyleFormats(XlsxPivotTable pivotTable) : IXlsxPivotTableStyleFormats
{
    private readonly XlsxPivotTable _pivotTable = pivotTable;

    #region IXlsxPivotTableStyleFormats members

    public IXlsxPivotStyleFormats ColumnGrandTotalFormats => new XlsxPivotStyleFormats(_pivotTable, isRowGrand: false);

    public IXlsxPivotStyleFormats RowGrandTotalFormats => new XlsxPivotStyleFormats(_pivotTable, isRowGrand: true);

    #endregion IXlsxPivotTableStyleFormats members
}
