#nullable disable

using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.ConditionalFormats
{
    public interface IXlsxCFColorScaleMid
    {
        IXlsxCFColorScaleMax Midpoint(XlsxCFContentType type, String value, XlsxColor color);
        IXlsxCFColorScaleMax Midpoint(XlsxCFContentType type, Double value, XlsxColor color);
        void Maximum(XlsxCFContentType type, String value, XlsxColor color);
        void Maximum(XlsxCFContentType type, Double value, XlsxColor color);
        void HighestValue(XlsxColor color);
    }
}
