#nullable disable

namespace Aid.Xlsx.ConditionalFormats
{
    public interface IXlsxCFDataBarMax
    {
        void Maximum(XlsxCFContentType type, String value);
        void Maximum(XlsxCFContentType type, Double value);
        void HighestValue();
    }
}
