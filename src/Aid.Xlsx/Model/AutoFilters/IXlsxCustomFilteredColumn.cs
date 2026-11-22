using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.AutoFilters;

public interface IXlsxCustomFilteredColumn
{
    void EqualTo(XlsxCellValue value, bool reapply = true);
    void NotEqualTo(XlsxCellValue value, bool reapply = true);
    void GreaterThan(XlsxCellValue value, bool reapply = true);
    void LessThan(XlsxCellValue value, bool reapply = true);
    void EqualOrGreaterThan(XlsxCellValue value, bool reapply = true);
    void EqualOrLessThan(XlsxCellValue value, bool reapply = true);
    void BeginsWith(String value, bool reapply = true);
    void NotBeginsWith(String value, bool reapply = true);
    void EndsWith(String value, bool reapply = true);
    void NotEndsWith(String value, bool reapply = true);
    void Contains(String value, bool reapply = true);
    void NotContains(String value, bool reapply = true);
}
