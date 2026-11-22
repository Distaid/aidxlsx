using Aid.Xlsx.Cells;

namespace Aid.Xlsx.AutoFilters;

internal class XlsxCustomFilteredColumn : IXlsxCustomFilteredColumn
{
    private readonly XlsxFilterColumn _filterColumn;
    private readonly XlsxConnector _connector;

    public XlsxCustomFilteredColumn(XlsxFilterColumn filterColumn, XlsxConnector connector)
    {
        _filterColumn = filterColumn;
        _connector = connector;
    }

    public void EqualTo(XlsxCellValue value, Boolean reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.Equal, reapply);
    }

    public void NotEqualTo(XlsxCellValue value, Boolean reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.NotEqual, reapply);
    }

    public void GreaterThan(XlsxCellValue value, Boolean reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.GreaterThan, reapply);
    }

    public void LessThan(XlsxCellValue value, Boolean reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.LessThan, reapply);
    }

    public void EqualOrGreaterThan(XlsxCellValue value, Boolean reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.EqualOrGreaterThan, reapply);
    }

    public void EqualOrLessThan(XlsxCellValue value, Boolean reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.EqualOrLessThan, reapply);
    }

    public void BeginsWith(String value, Boolean reapply)
    {
        ApplyWildcardCustomFilter(value + "*", true, reapply);
    }

    public void NotBeginsWith(String value, Boolean reapply)
    {
        ApplyWildcardCustomFilter(value + "*", false, reapply);
    }

    public void EndsWith(String value, Boolean reapply)
    {
        ApplyWildcardCustomFilter("*" + value, true, reapply);
    }

    public void NotEndsWith(String value, Boolean reapply)
    {
        ApplyWildcardCustomFilter("*" + value, false, reapply);
    }

    public void Contains(String value, Boolean reapply)
    {
        ApplyWildcardCustomFilter("*" + value + "*", true, reapply);
    }

    public void NotContains(String value, Boolean reapply)
    {
        ApplyWildcardCustomFilter("*" + value + "*", false, reapply);
    }

    private void ApplyCustomFilter(XlsxCellValue value, XlsxFilterOperator op, bool reapply)
    {
        _filterColumn.AddFilter(XlsxFilter.CreateCustomFilter(value, op, _connector), reapply);
    }

    private void ApplyWildcardCustomFilter(string pattern, bool match, bool reapply)
    {
        _filterColumn.AddFilter(XlsxFilter.CreateCustomPatternFilter(pattern, match, _connector), reapply);
    }
}
