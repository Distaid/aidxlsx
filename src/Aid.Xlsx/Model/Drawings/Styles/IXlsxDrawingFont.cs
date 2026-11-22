#nullable disable

using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.Drawings.Styles
{
    public interface IXlsxDrawingFont : IXlsxFontBase
    {
        IXlsxDrawingStyle SetBold(); IXlsxDrawingStyle SetBold(Boolean value);
        IXlsxDrawingStyle SetItalic(); IXlsxDrawingStyle SetItalic(Boolean value);
        IXlsxDrawingStyle SetUnderline(); IXlsxDrawingStyle SetUnderline(XlsxFontUnderlineValues value);
        IXlsxDrawingStyle SetStrikethrough(); IXlsxDrawingStyle SetStrikethrough(Boolean value);
        IXlsxDrawingStyle SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value);
        IXlsxDrawingStyle SetShadow(); IXlsxDrawingStyle SetShadow(Boolean value);
        IXlsxDrawingStyle SetFontSize(Double value);
        IXlsxDrawingStyle SetFontColor(XlsxColor value);
        IXlsxDrawingStyle SetFontName(String value);
        IXlsxDrawingStyle SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value);
        IXlsxDrawingStyle SetFontCharSet(XlsxFontCharSet value);
        IXlsxDrawingStyle SetFontScheme(XlsxFontScheme value);
    }
}
