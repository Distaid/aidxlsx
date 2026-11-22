#nullable disable

using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.ConditionalFormats
{
    public interface IXlsxCFColorScaleMax
    {
        void Maximum(XlsxCFContentType type, String value, XlsxColor color);
        void Maximum(XlsxCFContentType type, Double value, XlsxColor color);
        void HighestValue(XlsxColor color);
    }
}
