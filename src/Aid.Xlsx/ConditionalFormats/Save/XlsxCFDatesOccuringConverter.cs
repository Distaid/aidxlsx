using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFDatesOccurringConverter : IXlsxCFConverter
    {
        private static readonly IDictionary<XlsxTimePeriod, string> formulaTemplates = new Dictionary<XlsxTimePeriod, string>()
        {
            [XlsxTimePeriod.Today] = "FLOOR({0},1)=TODAY()",
            [XlsxTimePeriod.Yesterday] = "FLOOR({0},1)=TODAY()-1",
            [XlsxTimePeriod.Tomorrow] = "FLOOR({0},1)=TODAY()+1",
            [XlsxTimePeriod.InTheLast7Days] = "AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())",
            [XlsxTimePeriod.ThisMonth] = "AND(MONTH({0})=MONTH(TODAY()),YEAR({0})=YEAR(TODAY()))",
            [XlsxTimePeriod.LastMonth] = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))",
            [XlsxTimePeriod.NextMonth] = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)),YEAR({0})=YEAR(EDATE(TODAY(),0+1)))",
            [XlsxTimePeriod.ThisWeek] = "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))",
            [XlsxTimePeriod.LastWeek] = "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))",
            [XlsxTimePeriod.NextWeek] = "AND(ROUNDDOWN({0},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({0},0)-TODAY()<(15-WEEKDAY(TODAY())))"
        };

        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            conditionalFormattingRule.TimePeriod = cf.TimePeriod.ToOpenXml();

            var address = cf.Range.RangeAddress.FirstAddress.ToStringRelative(false);
            var formula = new Formula { Text = String.Format(formulaTemplates[cf.TimePeriod], address) };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
