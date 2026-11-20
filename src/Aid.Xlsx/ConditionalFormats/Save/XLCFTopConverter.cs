using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Aid.Xlsx
{
    internal class XLCFTopConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(XLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            UInt32 val = UInt32.Parse(cf.Values[1].Value);
            var conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(cf, priority, context);
            conditionalFormattingRule.Percent = cf.Percent;
            conditionalFormattingRule.Rank = val;
            conditionalFormattingRule.Bottom = cf.Bottom;
            return conditionalFormattingRule;
        }
    }
}
