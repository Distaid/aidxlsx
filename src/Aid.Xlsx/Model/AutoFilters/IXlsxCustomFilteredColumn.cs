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
    void BeginsWith(string value, bool reapply = true);
    void NotBeginsWith(string value, bool reapply = true);
    void EndsWith(string value, bool reapply = true);
    void NotEndsWith(string value, bool reapply = true);
    void Contains(string value, bool reapply = true);
    void NotContains(string value, bool reapply = true);
}
