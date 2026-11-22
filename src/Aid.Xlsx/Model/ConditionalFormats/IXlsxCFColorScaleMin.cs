#nullable disable

using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.ConditionalFormats
{
    public enum XlsxCFContentType { Number, Percent, Formula, Percentile, Minimum, Maximum }
    public interface IXlsxCFColorScaleMin
    {
        IXlsxCFColorScaleMid Minimum(XlsxCFContentType type, String value, XlsxColor color);
        IXlsxCFColorScaleMid Minimum(XlsxCFContentType type, Double value, XlsxColor color);
        IXlsxCFColorScaleMid LowestValue(XlsxColor color);
    }
}
