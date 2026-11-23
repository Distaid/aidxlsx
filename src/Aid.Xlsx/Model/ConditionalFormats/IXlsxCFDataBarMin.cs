#nullable disable

namespace Aid.Xlsx.Model.ConditionalFormats;

public interface IXlsxCFDataBarMin
{
    IXlsxCFDataBarMax Minimum(XlsxCFContentType type, string value);
    IXlsxCFDataBarMax Minimum(XlsxCFContentType type, double value);
    IXlsxCFDataBarMax LowestValue();
}
