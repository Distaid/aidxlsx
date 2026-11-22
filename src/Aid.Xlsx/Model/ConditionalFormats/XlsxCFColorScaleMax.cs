#nullable disable

using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.ConditionalFormats
{
    internal class XlsxCFColorScaleMax : IXlsxCFColorScaleMax
    {
        private readonly XlsxConditionalFormat _conditionalFormat;
        public XlsxCFColorScaleMax(XlsxConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public void Maximum(XlsxCFContentType type, String value, XlsxColor color)
        {
            _conditionalFormat.Values.Add(new XlsxFormula { Value = value });
            _conditionalFormat.Colors.Add(color);
            _conditionalFormat.ContentTypes.Add(type);
        }
        public void Maximum(XlsxCFContentType type, Double value, XlsxColor color)
        {
            Maximum(type, value.ToInvariantString(), color);
        }
        public void HighestValue(XlsxColor color)
        {
            _conditionalFormat.Values.Add(null);
            _conditionalFormat.Colors.Add(color);
            _conditionalFormat.ContentTypes.Add(XlsxCFContentType.Maximum);
        }
    }
}
