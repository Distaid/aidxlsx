#nullable disable

namespace Aid.Xlsx.Model.ConditionalFormats;

public enum XlsxCFIconSetOperator { GreaterThan, EqualOrGreaterThan }
public interface IXlsxCFIconSet
{
    IXlsxCFIconSet AddValue(XlsxCFIconSetOperator setOperator, string value, XlsxCFContentType type);
    IXlsxCFIconSet AddValue(XlsxCFIconSetOperator setOperator, double value, XlsxCFContentType type);
}
