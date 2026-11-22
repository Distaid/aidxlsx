using Aid.Xlsx.Extensions;
using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFDataBarConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XlsxCFBaseConverter.Convert(cf, priority);

            var dataBar = new DataBar { ShowValue = !cf.ShowBarOnly };

            var conditionalFormatValueObject1 = GetConditionalFormatValueObjectByIndex(cf, 1, ConditionalFormatValueObjectValues.Min);
            var conditionalFormatValueObject2 = GetConditionalFormatValueObjectByIndex(cf, 2, ConditionalFormatValueObjectValues.Max);

            var color = new Color();
            switch (cf.Colors[1].ColorType)
            {
                case XlsxColorType.Color:
                    color.Rgb = cf.Colors[1].Color.ToHex();
                    break;

                case XlsxColorType.Theme:
                    color.Theme = System.Convert.ToUInt32(cf.Colors[1].ThemeColor);
                    break;

                case XlsxColorType.Indexed:
                    color.Indexed = System.Convert.ToUInt32(cf.Colors[1].Indexed);
                    break;
            }

            dataBar.Append(conditionalFormatValueObject1);
            dataBar.Append(conditionalFormatValueObject2);
            dataBar.Append(color);

            conditionalFormattingRule.Append(dataBar);

            var conditionalFormattingRuleExtensionList = new ConditionalFormattingRuleExtensionList();
            conditionalFormattingRuleExtensionList.Append(BuildRuleExtension(cf));
            conditionalFormattingRule.Append(conditionalFormattingRuleExtensionList);

            return conditionalFormattingRule;
        }

        private ConditionalFormattingRuleExtension BuildRuleExtension(IXlsxConditionalFormat cf)
        {
            var conditionalFormattingRuleExtension = new ConditionalFormattingRuleExtension { Uri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" };
            conditionalFormattingRuleExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var id = new DocumentFormat.OpenXml.Office2010.Excel.Id
            {
                Text = ((XlsxConditionalFormat)cf).Id.WrapInBraces()
            };
            conditionalFormattingRuleExtension.Append(id);

            return conditionalFormattingRuleExtension;
        }

        private ConditionalFormatValueObject GetConditionalFormatValueObjectByIndex(IXlsxConditionalFormat cf, int index, ConditionalFormatValueObjectValues defaultType)
        {
            var conditionalFormatValueObject = new ConditionalFormatValueObject();

            if (cf.ContentTypes.TryGetValue(index, out var contentType))
            {
                conditionalFormatValueObject.Type = contentType.ToOpenXml();
            }
            else
            {
                conditionalFormatValueObject.Type = defaultType;
            }

            if (cf.Values.TryGetValue(index, out var value1) && value1?.Value != null)
            {
                conditionalFormatValueObject.Val = value1.Value;
            }

            return conditionalFormatValueObject;
        }
    }
}
