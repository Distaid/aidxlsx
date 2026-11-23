#nullable disable

namespace Aid.Xlsx.Model.ConditionalFormats;

public interface IXlsxCFDataBarMax
{
    void Maximum(XlsxCFContentType type, string value);
    void Maximum(XlsxCFContentType type, double value);
    void HighestValue();
}
