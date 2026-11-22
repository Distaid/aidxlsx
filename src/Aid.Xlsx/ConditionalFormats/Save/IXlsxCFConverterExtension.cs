#nullable disable

using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace Aid.Xlsx.ConditionalFormats.Save
{
    internal interface IXlsxCFConverterExtension
    {
        ConditionalFormattingRule Convert(IXlsxConditionalFormat cf, XlsxWorkbook.SaveContext context);
    }
}
