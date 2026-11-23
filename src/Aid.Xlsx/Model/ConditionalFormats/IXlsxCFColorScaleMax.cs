#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.ConditionalFormats;

public interface IXlsxCFColorScaleMax
{
    void Maximum(XlsxCFContentType type, string value, XlsxColor color);
    void Maximum(XlsxCFContentType type, double value, XlsxColor color);
    void HighestValue(XlsxColor color);
}
