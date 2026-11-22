#nullable disable

using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.RichText
{
    public interface IXlsxFormattedText<T> : IEnumerable<IXlsxRichString>, IEquatable<IXlsxFormattedText<T>>, IXlsxWithRichString
    {
        Boolean Bold { set; }
        Boolean Italic { set; }
        XlsxFontUnderlineValues Underline { set; }
        Boolean Strikethrough { set; }
        XlsxFontVerticalTextAlignmentValues VerticalAlignment { set; }
        Boolean Shadow { set; }
        Double FontSize { set; }
        XlsxColor FontColor { set; }
        String FontName { set; }
        XlsxFontFamilyNumberingValues FontFamilyNumbering { set; }

        IXlsxFormattedText<T> SetBold(); IXlsxFormattedText<T> SetBold(Boolean value);
        IXlsxFormattedText<T> SetItalic(); IXlsxFormattedText<T> SetItalic(Boolean value);
        IXlsxFormattedText<T> SetUnderline(); IXlsxFormattedText<T> SetUnderline(XlsxFontUnderlineValues value);
        IXlsxFormattedText<T> SetStrikethrough(); IXlsxFormattedText<T> SetStrikethrough(Boolean value);
        IXlsxFormattedText<T> SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value);
        IXlsxFormattedText<T> SetShadow(); IXlsxFormattedText<T> SetShadow(Boolean value);
        IXlsxFormattedText<T> SetFontSize(Double value);
        IXlsxFormattedText<T> SetFontColor(XlsxColor value);
        IXlsxFormattedText<T> SetFontName(String value);
        IXlsxFormattedText<T> SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value);

        IXlsxRichString AddText(String text, IXlsxFontBase font);
        IXlsxFormattedText<T> ClearText();
        IXlsxFormattedText<T> ClearFont();

        IXlsxFormattedText<T> Substring(Int32 index);
        IXlsxFormattedText<T> Substring(Int32 index, Int32 length);

        /// <summary>
        /// Replace the text and formatting of this text by texts and formatting from the <paramref name="original"/> text.
        /// </summary>
        /// <param name="original">Original to copy from.</param>
        /// <returns>This text.</returns>
        IXlsxFormattedText<T> CopyFrom(IXlsxFormattedText<T> original);

        /// <summary>
        /// How many rich strings is the formatted text composed of.
        /// </summary>
        Int32 Count { get; }

        /// <summary>
        /// Length of the whole formatted text.
        /// </summary>
        Int32 Length { get; }

        /// <summary>
        /// Get text of the whole formatted text.
        /// </summary>
        String Text { get; }

        /// <summary>
        /// Does this text has phonetics? Unlike accessing the <see cref="Phonetics"/> property, this method
        /// doesn't create a new instance on access.
        /// </summary>
        Boolean HasPhonetics { get; }

        /// <summary>
        /// Get or create phonetics for the text. Use <see cref="HasPhonetics"/> to check for existence to avoid unnecessary creation.
        /// </summary>
        IXlsxPhonetics Phonetics { get; }
    }
}
