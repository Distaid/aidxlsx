using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Aid.Xlsx
{
    internal class XLCFNotErrorConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(XLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(cf, priority, context);
            var formula = new Formula { Text = "NOT(ISERROR(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
