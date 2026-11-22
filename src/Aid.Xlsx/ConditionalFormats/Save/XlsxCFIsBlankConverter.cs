using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFIsBlankConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            var formula = new Formula { Text = "LEN(TRIM(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))=0" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
