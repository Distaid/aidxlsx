using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFTopConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            UInt32 val = UInt32.Parse(cf.Values[1].Value);
            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            conditionalFormattingRule.Percent = cf.Percent;
            conditionalFormattingRule.Rank = val;
            conditionalFormattingRule.Bottom = cf.Bottom;
            return conditionalFormattingRule;
        }
    }
}
