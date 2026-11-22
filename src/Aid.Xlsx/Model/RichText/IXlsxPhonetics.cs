using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.RichText;

public enum XlsxPhoneticAlignment { Center = 0, Distributed = 1, Left = 2, NoControl = 3 }
public enum XlsxPhoneticType { FullWidthKatakana = 0, HalfWidthKatakana = 1, Hiragana = 2, NoConversion = 3 }
public interface IXlsxPhonetics : IXlsxFontBase, IEnumerable<IXlsxPhonetic>, IEquatable<IXlsxPhonetics>
{
    IXlsxPhonetics SetBold(); IXlsxPhonetics SetBold(Boolean value);
    IXlsxPhonetics SetItalic(); IXlsxPhonetics SetItalic(Boolean value);
    IXlsxPhonetics SetUnderline(); IXlsxPhonetics SetUnderline(XlsxFontUnderlineValues value);
    IXlsxPhonetics SetStrikethrough(); IXlsxPhonetics SetStrikethrough(Boolean value);
    IXlsxPhonetics SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value);
    IXlsxPhonetics SetShadow(); IXlsxPhonetics SetShadow(Boolean value);
    IXlsxPhonetics SetFontSize(Double value);
    IXlsxPhonetics SetFontColor(XlsxColor value);
    IXlsxPhonetics SetFontName(String value);
    IXlsxPhonetics SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value);
    IXlsxPhonetics SetFontCharSet(XlsxFontCharSet value);
    IXlsxPhonetics SetFontScheme(XlsxFontScheme value);

    /// <summary>
    /// Add a phonetic run above a base text. Phonetic runs can't overlap.
    /// </summary>
    /// <param name="text">Text to display above a section of a base text. Can't be empty.</param>
    /// <param name="start">Index of a first character of a base  text above which should <paramref name="text"/> be displayed. Valid values are <c>0</c>..<c>length-1</c>.</param>
    /// <param name="end">The excluded ending index in a base text (the hint is not displayed above the <c>end</c>). Must be &gt; <paramref name="start"/>. Valid values are <c>1</c>..<c>length</c>.</param>
    IXlsxPhonetics Add(String text, Int32 start, Int32 end);

    /// <summary>
    /// Remove all phonetic runs. Keeps font properties.
    /// </summary>
    IXlsxPhonetics ClearText();

    /// <summary>
    /// Reset font properties to the default font of a container (likely <c>IXlsxCell</c>). Keeps phonetic runs, <see cref="Type"/> and <see cref="Alignment"/>.
    /// </summary>
    IXlsxPhonetics ClearFont();

    /// <summary>
    /// Number of phonetic runs above the base text.
    /// </summary>
    Int32 Count { get; }

    XlsxPhoneticAlignment Alignment { get; set; }
    XlsxPhoneticType Type { get; set; }

    IXlsxPhonetics SetAlignment(XlsxPhoneticAlignment phoneticAlignment);
    IXlsxPhonetics SetType(XlsxPhoneticType phoneticType);
}
