#nullable disable

using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Extensions;

namespace Aid.Xlsx.ConditionalFormats
{
    internal class XlsxCFDataBarMax : IXlsxCFDataBarMax
    {
        private readonly XlsxConditionalFormat _conditionalFormat;
        public XlsxCFDataBarMax(XlsxConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public void Maximum(XlsxCFContentType type, String value)
        {
            _conditionalFormat.ContentTypes.Add(type);
            _conditionalFormat.Values.Add(new XlsxFormula { Value = value });
        }
        public void Maximum(XlsxCFContentType type, Double value)
        {
            Maximum(type, value.ToInvariantString());
        }

        public void HighestValue()
        {
            Maximum(XlsxCFContentType.Maximum, "0");
        }
    }
}
