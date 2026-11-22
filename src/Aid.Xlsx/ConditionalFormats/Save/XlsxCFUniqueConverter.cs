using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFUniqueConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            return conditionalFormattingRule;
        }
    }
}
