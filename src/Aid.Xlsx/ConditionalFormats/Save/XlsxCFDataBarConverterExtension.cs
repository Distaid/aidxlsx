using Aid.Xlsx.Extensions;
using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Office.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFDataBarConverterExtension : IXlsxCFConverterExtension
    {
        private static readonly IReadOnlyDictionary<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues, ConditionalFormattingValueObjectTypeValues> CFValueToTypeMap =
            new Dictionary<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues, ConditionalFormattingValueObjectTypeValues>
            {
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Max, ConditionalFormattingValueObjectTypeValues.AutoMax },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Min, ConditionalFormattingValueObjectTypeValues.AutoMin },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Number, ConditionalFormattingValueObjectTypeValues.Numeric },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Percent, ConditionalFormattingValueObjectTypeValues.Percent },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Percentile, ConditionalFormattingValueObjectTypeValues.Percentile },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Formula, ConditionalFormattingValueObjectTypeValues.Formula },
            };

        public ConditionalFormattingRule Convert(IXlsxConditionalFormat cf, XlsxWorkbook.SaveContext context)
        {
            ConditionalFormattingRule conditionalFormattingRule = new ConditionalFormattingRule()
            {
                Type = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValues.DataBar,
                Id = ((XlsxConditionalFormat)cf).Id.WrapInBraces()
            };

            DataBar dataBar = new DataBar()
            {
                MinLength = 0,
                MaxLength = 100,
                Gradient = true,
                ShowValue = !cf.ShowBarOnly
            };

            var cfMinType = cf.ContentTypes.TryGetValue(1, out var contentType1)
                ? GetCFType(contentType1.ToOpenXml())
                : ConditionalFormattingValueObjectTypeValues.AutoMin;
            var cfMin = new ConditionalFormattingValueObject { Type = cfMinType };
            if (cf.Values.Any() && cf.Values[1]?.Value != null)
            {
                cfMin.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
                cfMin.Append(new Formula() { Text = cf.Values[1].Value });
            }

            var cfMaxType = cf.ContentTypes.TryGetValue(2, out var contentType2)
                ? GetCFType(contentType2.ToOpenXml())
                : ConditionalFormattingValueObjectTypeValues.AutoMax;
            var cfMax = new ConditionalFormattingValueObject { Type = cfMaxType };
            if (cf.Values.Count >= 2 && cf.Values[2]?.Value != null)
            {
                cfMax.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
                cfMax.Append(new Formula() { Text = cf.Values[2].Value });
            }

            var barAxisColor = new BarAxisColor { Rgb = XlsxColor.Black.Color.ToHex() };

            var negativeFillColor = new NegativeFillColor { Rgb = cf.Colors[1].Color.ToHex() };
            if (cf.Colors.Count == 2)
            {
                negativeFillColor = new NegativeFillColor { Rgb = cf.Colors[2].Color.ToHex() };
            }

            dataBar.Append(cfMin);
            dataBar.Append(cfMax);

            dataBar.Append(negativeFillColor);
            dataBar.Append(barAxisColor);

            conditionalFormattingRule.Append(dataBar);

            return conditionalFormattingRule;
        }

        private static ConditionalFormattingValueObjectTypeValues GetCFType(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues value)
        {
            return CFValueToTypeMap[value];
        }
    }
}
