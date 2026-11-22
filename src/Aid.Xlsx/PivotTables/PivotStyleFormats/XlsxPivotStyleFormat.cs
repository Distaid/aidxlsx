namespace Aid.Xlsx.PivotTables.PivotStyleFormats;

internal class XlsxPivotStyleFormat : XlsxPivotStyleFormatBase
{
    private readonly Func<XlsxPivotArea, bool> _filter;
    private readonly Func<XlsxPivotArea> _factory;

    public XlsxPivotStyleFormat(XlsxPivotTable pivotTable, Func<XlsxPivotArea, bool> filter, Func<XlsxPivotArea> factory)
        : base(pivotTable)
    {
        _filter = filter;
        _factory = factory;
    }

    internal override XlsxPivotArea GetCurrentArea()
    {
        return _factory();
    }

    internal override bool Filter(XlsxPivotArea area)
    {
        return _filter(area);
    }
}
