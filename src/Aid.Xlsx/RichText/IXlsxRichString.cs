#nullable disable

using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.RichText
{
    public interface IXlsxWithRichString
    {
        IXlsxRichString AddText(String text);
        IXlsxRichString AddNewLine();
    }
    public interface IXlsxRichString : IXlsxFontBase, IEquatable<IXlsxRichString>, IXlsxWithRichString
    {
        String Text { get; set; }

        IXlsxRichString SetBold(); IXlsxRichString SetBold(Boolean value);
        IXlsxRichString SetItalic(); IXlsxRichString SetItalic(Boolean value);
        IXlsxRichString SetUnderline(); IXlsxRichString SetUnderline(XlsxFontUnderlineValues value);
        IXlsxRichString SetStrikethrough(); IXlsxRichString SetStrikethrough(Boolean value);
        IXlsxRichString SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value);
        IXlsxRichString SetShadow(); IXlsxRichString SetShadow(Boolean value);
        IXlsxRichString SetFontSize(Double value);
        IXlsxRichString SetFontColor(XlsxColor value);
        IXlsxRichString SetFontName(String value);
        IXlsxRichString SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value);
        IXlsxRichString SetFontCharSet(XlsxFontCharSet value);

        /// <inheritdoc cref="IXlsxFontBase.FontScheme"/>
        IXlsxRichString SetFontScheme(XlsxFontScheme value);
    }
}
