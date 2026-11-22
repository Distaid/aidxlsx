#nullable disable

namespace Aid.Xlsx.Model.ConditionalFormats;

public enum XlsxCFIconSetOperator { GreaterThan, EqualOrGreaterThan }
public interface IXlsxCFIconSet
{
    IXlsxCFIconSet AddValue(XlsxCFIconSetOperator setOperator, String value, XlsxCFContentType type);
    IXlsxCFIconSet AddValue(XlsxCFIconSetOperator setOperator, Double value, XlsxCFContentType type);
}
