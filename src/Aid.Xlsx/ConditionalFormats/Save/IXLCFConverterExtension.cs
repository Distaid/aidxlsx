#nullable disable

using DocumentFormat.OpenXml.Office2010.Excel;

namespace Aid.Xlsx
{
    internal interface IXLCFConverterExtension
    {
        ConditionalFormattingRule Convert(IXLConditionalFormat cf, XLWorkbook.SaveContext context);
    }
}
