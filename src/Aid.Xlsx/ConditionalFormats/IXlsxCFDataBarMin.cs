#nullable disable

namespace Aid.Xlsx.ConditionalFormats
{
    public interface IXlsxCFDataBarMin
    {
        IXlsxCFDataBarMax Minimum(XlsxCFContentType type, String value);
        IXlsxCFDataBarMax Minimum(XlsxCFContentType type, Double value);
        IXlsxCFDataBarMax LowestValue();
    }
}
