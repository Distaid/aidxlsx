using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Infrastructure.Caching;

internal class XlsxRangeRepository : XlsxWorkbookElementRepositoryBase<XlsxRangeKey, XlsxRangeBase>
{
    public XlsxRangeRepository(XlsxWorkbook workbook, Func<XlsxRangeKey, XlsxRangeBase> createNew) : base(workbook, createNew)
    {
    }

    public XlsxRangeRepository(XlsxWorkbook workbook, Func<XlsxRangeKey, XlsxRangeBase> createNew, IEqualityComparer<XlsxRangeKey> comparer) : base(workbook, createNew, comparer)
    {
    }
}
