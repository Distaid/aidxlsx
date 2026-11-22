#nullable disable

namespace Aid.Xlsx.Model.ConditionalFormats;

public interface IXlsxCFDataBarMin
{
    IXlsxCFDataBarMax Minimum(XlsxCFContentType type, String value);
    IXlsxCFDataBarMax Minimum(XlsxCFContentType type, Double value);
    IXlsxCFDataBarMax LowestValue();
}
