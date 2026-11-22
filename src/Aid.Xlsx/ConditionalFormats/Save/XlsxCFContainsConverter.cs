using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFContainsConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            String val = cf.Values[1].Value;
            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.ContainsText;
            conditionalFormattingRule.Text = val;

            var formula = new Formula { Text = "NOT(ISERROR(SEARCH(\"" + val + "\"," + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")))" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
