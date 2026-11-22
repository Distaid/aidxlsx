using Aid.Xlsx.Ranges;
using Aid.Xlsx.Workbooks;

namespace Aid.Xlsx.Caching
{
    internal class XlsxRangeRepository : XlsxWorkbookElementRepositoryBase<XlsxRangeKey, XlsxRangeBase>
    {
        public XlsxRangeRepository(XlsxWorkbook workbook, Func<XlsxRangeKey, XlsxRangeBase> createNew) : base(workbook, createNew)
        {
        }

        public XlsxRangeRepository(XlsxWorkbook workbook, Func<XlsxRangeKey, XlsxRangeBase> createNew, IEqualityComparer<XlsxRangeKey> comparer) : base(workbook, createNew, comparer)
        {
        }
    }
}
