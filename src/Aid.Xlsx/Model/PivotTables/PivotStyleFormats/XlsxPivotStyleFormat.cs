namespace Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

internal class XlsxPivotStyleFormat(XlsxPivotTable pivotTable, Func<XlsxPivotArea, bool> filter, Func<XlsxPivotArea> factory) : XlsxPivotStyleFormatBase(pivotTable)
{
    private readonly Func<XlsxPivotArea, bool> _filter = filter;
    private readonly Func<XlsxPivotArea> _factory = factory;

    internal override XlsxPivotArea GetCurrentArea()
    {
        return _factory();
    }

    internal override bool Filter(XlsxPivotArea area)
    {
        return _filter(area);
    }
}
