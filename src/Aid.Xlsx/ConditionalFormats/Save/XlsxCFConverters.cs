#nullable disable

using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFConverters
    {
        private static readonly Dictionary<XlsxConditionalFormatType, IXlsxCFConverter> Converters;
        static XlsxCFConverters()
        {
            Converters = new Dictionary<XlsxConditionalFormatType, IXlsxCFConverter>
            {
                {XlsxConditionalFormatType.ColorScale, new XlsxCFColorScaleConverter()},
                {XlsxConditionalFormatType.StartsWith, new XlsxCFStartsWithConverter()},
                {XlsxConditionalFormatType.EndsWith, new XlsxCFEndsWithConverter()},
                {XlsxConditionalFormatType.IsBlank, new XlsxCFIsBlankConverter()},
                {XlsxConditionalFormatType.NotBlank, new XlsxCFNotBlankConverter()},
                {XlsxConditionalFormatType.IsError, new XlsxCFIsErrorConverter()},
                {XlsxConditionalFormatType.NotError, new XlsxCFNotErrorConverter()},
                {XlsxConditionalFormatType.ContainsText, new XlsxCFContainsConverter()},
                {XlsxConditionalFormatType.NotContainsText, new XlsxCFNotContainsConverter()},
                {XlsxConditionalFormatType.CellIs, new XlsxCFCellIsConverter()},
                {XlsxConditionalFormatType.IsUnique, new XlsxCFUniqueConverter()},
                {XlsxConditionalFormatType.IsDuplicate, new XlsxCFUniqueConverter()},
                {XlsxConditionalFormatType.Expression, new XlsxCFCellIsConverter()},
                {XlsxConditionalFormatType.Top10, new XlsxCFTopConverter()},
                {XlsxConditionalFormatType.DataBar, new XlsxCFDataBarConverter()},
                {XlsxConditionalFormatType.IconSet, new XlsxCFIconSetConverter()},
                {XlsxConditionalFormatType.TimePeriod, new XlsxCFDatesOccurringConverter()}
            };
        }

        public static ConditionalFormattingRule Convert(XlsxConditionalFormat conditionalFormat, Int32 priority, XlsxWorkbook.SaveContext context)
        {
            if (!Converters.TryGetValue(conditionalFormat.ConditionalFormatType, out var converter))
                throw new NotImplementedException(string.Format("Conditional formatting rule '{0}' hasn't been implemented", conditionalFormat.ConditionalFormatType));

            return converter.Convert(conditionalFormat, priority, context);
        }
    }
}
