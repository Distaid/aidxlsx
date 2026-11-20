using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Aid.Xlsx
{
    internal class XLCFUniqueConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(XLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(cf, priority, context);
            return conditionalFormattingRule;
        }
    }
}
