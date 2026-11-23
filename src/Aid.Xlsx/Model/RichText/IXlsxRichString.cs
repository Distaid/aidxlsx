#nullable disable

using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.RichText;

public interface IXlsxWithRichString
{
    IXlsxRichString AddText(string text);
    IXlsxRichString AddNewLine();
}
public interface IXlsxRichString : IXlsxFontBase, IEquatable<IXlsxRichString>, IXlsxWithRichString
{
    string Text { get; set; }

    IXlsxRichString SetBold(); IXlsxRichString SetBold(bool value);
    IXlsxRichString SetItalic(); IXlsxRichString SetItalic(bool value);
    IXlsxRichString SetUnderline(); IXlsxRichString SetUnderline(XlsxFontUnderlineValues value);
    IXlsxRichString SetStrikethrough(); IXlsxRichString SetStrikethrough(bool value);
    IXlsxRichString SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value);
    IXlsxRichString SetShadow(); IXlsxRichString SetShadow(bool value);
    IXlsxRichString SetFontSize(double value);
    IXlsxRichString SetFontColor(XlsxColor value);
    IXlsxRichString SetFontName(string value);
    IXlsxRichString SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value);
    IXlsxRichString SetFontCharSet(XlsxFontCharSet value);

    /// <inheritdoc cref="IXlsxFontBase.FontScheme"/>
    IXlsxRichString SetFontScheme(XlsxFontScheme value);
}
