#nullable disable

using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.RichText;

public interface IXlsxFormattedText<T> : IEnumerable<IXlsxRichString>, IEquatable<IXlsxFormattedText<T>>, IXlsxWithRichString
{
    bool Bold { set; }
    bool Italic { set; }
    XlsxFontUnderlineValues Underline { set; }
    bool Strikethrough { set; }
    XlsxFontVerticalTextAlignmentValues VerticalAlignment { set; }
    bool Shadow { set; }
    double FontSize { set; }
    XlsxColor FontColor { set; }
    string FontName { set; }
    XlsxFontFamilyNumberingValues FontFamilyNumbering { set; }

    IXlsxFormattedText<T> SetBold(); IXlsxFormattedText<T> SetBold(bool value);
    IXlsxFormattedText<T> SetItalic(); IXlsxFormattedText<T> SetItalic(bool value);
    IXlsxFormattedText<T> SetUnderline(); IXlsxFormattedText<T> SetUnderline(XlsxFontUnderlineValues value);
    IXlsxFormattedText<T> SetStrikethrough(); IXlsxFormattedText<T> SetStrikethrough(bool value);
    IXlsxFormattedText<T> SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value);
    IXlsxFormattedText<T> SetShadow(); IXlsxFormattedText<T> SetShadow(bool value);
    IXlsxFormattedText<T> SetFontSize(double value);
    IXlsxFormattedText<T> SetFontColor(XlsxColor value);
    IXlsxFormattedText<T> SetFontName(string value);
    IXlsxFormattedText<T> SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value);

    IXlsxRichString AddText(string text, IXlsxFontBase font);
    IXlsxFormattedText<T> ClearText();
    IXlsxFormattedText<T> ClearFont();

    IXlsxFormattedText<T> Substring(int index);
    IXlsxFormattedText<T> Substring(int index, int length);

    /// <summary>
    /// Replace the text and formatting of this text by texts and formatting from the <paramref name="original"/> text.
    /// </summary>
    /// <param name="original">Original to copy from.</param>
    /// <returns>This text.</returns>
    IXlsxFormattedText<T> CopyFrom(IXlsxFormattedText<T> original);

    /// <summary>
    /// How many rich strings is the formatted text composed of.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Length of the whole formatted text.
    /// </summary>
    int Length { get; }

    /// <summary>
    /// Get text of the whole formatted text.
    /// </summary>
    string Text { get; }

    /// <summary>
    /// Does this text has phonetics? Unlike accessing the <see cref="Phonetics"/> property, this method
    /// doesn't create a new instance on access.
    /// </summary>
    bool HasPhonetics { get; }

    /// <summary>
    /// Get or create phonetics for the text. Use <see cref="HasPhonetics"/> to check for existence to avoid unnecessary creation.
    /// </summary>
    IXlsxPhonetics Phonetics { get; }
}
