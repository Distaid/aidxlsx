#nullable disable

using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Infrastructure.Caching;

internal sealed class XlsxColorRepository : XlsxRepositoryBase<XlsxColorKey, XlsxColor>
{
    #region Constructors

    public XlsxColorRepository(Func<XlsxColorKey, XlsxColor> createNew)
        : base(createNew)
    {
    }

    public XlsxColorRepository(Func<XlsxColorKey, XlsxColor> createNew, IEqualityComparer<XlsxColorKey> comparer)
        : base(createNew, comparer)
    {
    }

    #endregion Constructors
}
