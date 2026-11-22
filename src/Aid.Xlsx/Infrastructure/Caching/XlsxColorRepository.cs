#nullable disable

using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.Caching
{
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
}
