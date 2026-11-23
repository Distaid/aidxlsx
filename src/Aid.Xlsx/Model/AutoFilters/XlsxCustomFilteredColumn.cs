using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.AutoFilters;

internal class XlsxCustomFilteredColumn(XlsxFilterColumn filterColumn, XlsxConnector connector) : IXlsxCustomFilteredColumn
{
    private readonly XlsxFilterColumn _filterColumn = filterColumn;
    private readonly XlsxConnector _connector = connector;

    public void EqualTo(XlsxCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.Equal, reapply);
    }

    public void NotEqualTo(XlsxCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.NotEqual, reapply);
    }

    public void GreaterThan(XlsxCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.GreaterThan, reapply);
    }

    public void LessThan(XlsxCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.LessThan, reapply);
    }

    public void EqualOrGreaterThan(XlsxCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.EqualOrGreaterThan, reapply);
    }

    public void EqualOrLessThan(XlsxCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XlsxFilterOperator.EqualOrLessThan, reapply);
    }

    public void BeginsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter(value + "*", true, reapply);
    }

    public void NotBeginsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter(value + "*", false, reapply);
    }

    public void EndsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter("*" + value, true, reapply);
    }

    public void NotEndsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter("*" + value, false, reapply);
    }

    public void Contains(string value, bool reapply)
    {
        ApplyWildcardCustomFilter("*" + value + "*", true, reapply);
    }

    public void NotContains(string value, bool reapply)
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
