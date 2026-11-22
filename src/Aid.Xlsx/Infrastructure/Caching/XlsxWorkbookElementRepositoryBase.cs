using Aid.Xlsx.Workbooks;

namespace Aid.Xlsx.Caching
{
    /// <summary>
    /// Base repository for <see cref="XlsxWorkbook"/> elements.
    /// </summary>
    internal abstract class XlsxWorkbookElementRepositoryBase<Tkey, Tvalue> : XlsxRepositoryBase<Tkey, Tvalue>
        where Tkey : struct, IEquatable<Tkey>
        where Tvalue : class
    {
        public XlsxWorkbook Workbook { get; private set; }

        public XlsxWorkbookElementRepositoryBase(XlsxWorkbook workbook, Func<Tkey, Tvalue> createNew)
            : this(workbook, createNew, EqualityComparer<Tkey>.Default)
        {
        }

        public XlsxWorkbookElementRepositoryBase(XlsxWorkbook workbook, Func<Tkey, Tvalue> createNew, IEqualityComparer<Tkey> comparer)
            : base(createNew, comparer)
        {
            Workbook = workbook;
        }
    }
}
