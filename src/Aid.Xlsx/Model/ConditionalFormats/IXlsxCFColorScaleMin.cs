#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.ConditionalFormats;

public enum XlsxCFContentType { Number, Percent, Formula, Percentile, Minimum, Maximum }
public interface IXlsxCFColorScaleMin
{
    IXlsxCFColorScaleMid Minimum(XlsxCFContentType type, string value, XlsxColor color);
    IXlsxCFColorScaleMid Minimum(XlsxCFContentType type, double value, XlsxColor color);
    IXlsxCFColorScaleMid LowestValue(XlsxColor color);
}
