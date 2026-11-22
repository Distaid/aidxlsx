using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal interface IXlsxCFConverter
    {
        ConditionalFormattingRule Convert(XlsxConditionalFormat cf, Int32 priority, XlsxWorkbook.SaveContext context);
    }
}
