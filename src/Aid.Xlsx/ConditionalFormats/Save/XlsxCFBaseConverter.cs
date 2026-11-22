using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal static class XlsxCFBaseConverter
    {
        public static ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority)
        {
            return new ConditionalFormattingRule
            {
                Type = cf.ConditionalFormatType.ToOpenXml(),
                Priority = priority,
                StopIfTrue = OpenXmlHelper.GetBooleanValue(cf.StopIfTrue, false)
            };
        }

        public static ConditionalFormattingRule ConvertWithDxf(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var cfRule = Convert(cf, priority);
            cfRule.FormatId = cf.FormatValue is not null ? context.GetDxfId(cf.FormatValue) : null;
            return cfRule;
        }
    }
}
