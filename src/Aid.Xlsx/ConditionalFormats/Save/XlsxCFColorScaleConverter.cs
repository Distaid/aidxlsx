#nullable disable

using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFColorScaleConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XlsxCFBaseConverter.Convert(cf, priority);

            var colorScale = new ColorScale();
            for (Int32 i = 1; i <= cf.ContentTypes.Count; i++)
            {
                var type = cf.ContentTypes[i].ToOpenXml();
                var val = cf.Values.TryGetValue(i, out XlsxFormula formula) ? formula?.Value : null;

                var conditionalFormatValueObject = new ConditionalFormatValueObject { Type = type };
                if (val != null)
                    conditionalFormatValueObject.Val = val;

                colorScale.Append(conditionalFormatValueObject);
            }

            for (Int32 i = 1; i <= cf.Colors.Count; i++)
            {
                var xlColor = cf.Colors[i];
                var color = new Color();
                switch (xlColor.ColorType)
                {
                    case XlsxColorType.Color:
                        color.Rgb = xlColor.Color.ToHex();
                        break;
                    case XlsxColorType.Theme:
                        color.Theme = System.Convert.ToUInt32(xlColor.ThemeColor);
                        break;

                    case XlsxColorType.Indexed:
                        color.Indexed = System.Convert.ToUInt32(xlColor.Indexed);
                        break;
                }

                colorScale.Append(color);
            }

            conditionalFormattingRule.Append(colorScale);

            return conditionalFormattingRule;
        }
    }
}
