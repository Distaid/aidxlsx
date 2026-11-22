using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFStartsWithConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            String? val = cf.Values[1].Value;
            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.BeginsWith;
            conditionalFormattingRule.Text = val;

            var formula = new Formula { Text = "LEFT(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "," + val.Length.ToString() + ")=\"" + val + "\"" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
