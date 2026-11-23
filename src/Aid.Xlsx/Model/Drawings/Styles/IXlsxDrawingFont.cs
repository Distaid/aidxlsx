#nullable disable

using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingFont : IXlsxFontBase
{
    IXlsxDrawingStyle SetBold(); IXlsxDrawingStyle SetBold(bool value);
    IXlsxDrawingStyle SetItalic(); IXlsxDrawingStyle SetItalic(bool value);
    IXlsxDrawingStyle SetUnderline(); IXlsxDrawingStyle SetUnderline(XlsxFontUnderlineValues value);
    IXlsxDrawingStyle SetStrikethrough(); IXlsxDrawingStyle SetStrikethrough(bool value);
    IXlsxDrawingStyle SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value);
    IXlsxDrawingStyle SetShadow(); IXlsxDrawingStyle SetShadow(bool value);
    IXlsxDrawingStyle SetFontSize(double value);
    IXlsxDrawingStyle SetFontColor(XlsxColor value);
    IXlsxDrawingStyle SetFontName(string value);
    IXlsxDrawingStyle SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value);
    IXlsxDrawingStyle SetFontCharSet(XlsxFontCharSet value);
    IXlsxDrawingStyle SetFontScheme(XlsxFontScheme value);
}
