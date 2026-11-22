#nullable disable

using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFIconSetConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XlsxCFBaseConverter.Convert(cf, priority);

            var iconSet = new IconSet { ShowValue = !cf.ShowIconOnly, Reverse = cf.ReverseIconOrder, IconSetValue = cf.IconSetStyle.ToOpenXml() };
            Int32 count = cf.Values.Count;
            for (Int32 i = 1; i <= count; i++)
            {
                var conditionalFormatValueObject = new ConditionalFormatValueObject { Type = cf.ContentTypes[i].ToOpenXml(), Val = cf.Values[i].Value, GreaterThanOrEqual = cf.IconSetOperators[i] == XlsxCFIconSetOperator.EqualOrGreaterThan };
                iconSet.Append(conditionalFormatValueObject);

            }
            conditionalFormattingRule.Append(iconSet);
            return conditionalFormattingRule;
        }
    }
}
