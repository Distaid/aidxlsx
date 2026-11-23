#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.ConditionalFormats;

public interface IXlsxCFColorScaleMid
{
    IXlsxCFColorScaleMax Midpoint(XlsxCFContentType type, string value, XlsxColor color);
    IXlsxCFColorScaleMax Midpoint(XlsxCFContentType type, double value, XlsxColor color);
    void Maximum(XlsxCFContentType type, string value, XlsxColor color);
    void Maximum(XlsxCFContentType type, double value, XlsxColor color);
    void HighestValue(XlsxColor color);
}
