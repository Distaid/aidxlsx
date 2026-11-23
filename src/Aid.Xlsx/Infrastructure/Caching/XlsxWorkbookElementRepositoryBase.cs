using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Infrastructure.Caching;

/// <summary>
/// Base repository for <see cref="XlsxWorkbook"/> elements.
/// </summary>
internal abstract class XlsxWorkbookElementRepositoryBase<Tkey, Tvalue>(XlsxWorkbook workbook, Func<Tkey, Tvalue> createNew, IEqualityComparer<Tkey> comparer) : XlsxRepositoryBase<Tkey, Tvalue>(createNew, comparer)
    where Tkey : struct, IEquatable<Tkey>
    where Tvalue : class
{
    public XlsxWorkbook Workbook { get; private set; } = workbook;

    public XlsxWorkbookElementRepositoryBase(XlsxWorkbook workbook, Func<Tkey, Tvalue> createNew)
        : this(workbook, createNew, EqualityComparer<Tkey>.Default)
    {
    }
}
