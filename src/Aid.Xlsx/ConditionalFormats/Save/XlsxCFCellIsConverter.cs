using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;
using Formula = DocumentFormat.OpenXml.Spreadsheet.Formula;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal class XlsxCFCellIsConverter : IXlsxCFConverter
    {
        public ConditionalFormattingRule Convert(XlsxConditionalFormat cf, int priority, XlsxWorkbook.SaveContext context)
        {
            String val = GetQuoted(cf.Values[1]);

            var conditionalFormattingRule = XlsxCFBaseConverter.ConvertWithDxf(cf, priority, context);
            conditionalFormattingRule.Operator = cf.Operator.ToOpenXml();

            var formula = new Formula(val);
            conditionalFormattingRule.Append(formula);

            if (cf.Operator == XlsxCFOperator.Between || cf.Operator == XlsxCFOperator.NotBetween)
            {
                var formula2 = new Formula { Text = GetQuoted(cf.Values[2]) };
                conditionalFormattingRule.Append(formula2);
            }

            return conditionalFormattingRule;
        }

        private String GetQuoted(XlsxFormula formula)
        {
            String value = formula.Value;

            if (formula.IsFormula ||
                value.StartsWith("\"") && value.EndsWith("\"") ||
                Double.TryParse(value, XlsxHelper.NumberStyle, XlsxHelper.ParseCulture, out _))
            {
                return value;
            }

            return String.Format("\"{0}\"", value.Replace("\"", "\"\""));
        }
    }
}
