using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFIsErrorConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            var formula = new Formula { Text = "ISERROR(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
