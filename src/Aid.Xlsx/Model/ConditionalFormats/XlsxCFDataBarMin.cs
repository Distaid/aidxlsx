#nullable disable

using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Extensions;

namespace Aid.Xlsx.ConditionalFormats
{
    internal class XlsxCFDataBarMin : IXlsxCFDataBarMin
    {
        private readonly XlsxConditionalFormat _conditionalFormat;
        public XlsxCFDataBarMin(XlsxConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public IXlsxCFDataBarMax Minimum(XlsxCFContentType type, String value)
        {
            _conditionalFormat.ContentTypes.Initialize(type);
            _conditionalFormat.Values.Initialize(new XlsxFormula { Value = value });
            return new XlsxCFDataBarMax(_conditionalFormat);
        }
        public IXlsxCFDataBarMax Minimum(XlsxCFContentType type, Double value)
        {
            return Minimum(type, value.ToInvariantString());
        }

        public IXlsxCFDataBarMax LowestValue()
        {
            return Minimum(XlsxCFContentType.Minimum, "0");
        }
    }
}
