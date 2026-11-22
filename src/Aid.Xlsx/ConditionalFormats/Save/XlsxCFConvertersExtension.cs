#nullable disable

using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFConvertersExtension
    {
        private readonly static Dictionary<XlsxConditionalFormatType, IXlsxCFConverterExtension> Converters;

        static XlsxCFConvertersExtension()
        {
            XlsxCFConvertersExtension.Converters = new Dictionary<XlsxConditionalFormatType, IXlsxCFConverterExtension>()
            {
                { XlsxConditionalFormatType.DataBar, new XlsxCFDataBarConverterExtension() }
            };
        }

        public XlsxCFConvertersExtension()
        {
        }

        public static ConditionalFormattingRule Convert(IXlsxConditionalFormat conditionalFormat, XlsxWorkbook.SaveContext context)
        {
            return XlsxCFConvertersExtension.Converters[conditionalFormat.ConditionalFormatType].Convert(conditionalFormat, context);
        }
    }
}
