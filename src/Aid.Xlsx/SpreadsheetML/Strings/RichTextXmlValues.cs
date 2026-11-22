using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Styles;

namespace Aid.Xlsx.SpreadsheetML.Strings;

internal static class RichTextXmlValues
{
    private static readonly string[] Underline = ["double", "doubleAccounting", "none", "single", "singleAccounting"];
    private static readonly string[] VerticalAlignment = ["baseline", "subscript", "superscript"];
    private static readonly string[] FontScheme = ["none", "major", "minor"];
    private static readonly string[] PhoneticAlignment = ["center", "distributed", "left", "noControl"];
    private static readonly string[] PhoneticType = ["fullwidthKatakana", "halfwidthKatakana", "Hiragana", "noConversion"];

    internal static string ToOpenXmlString(this XlsxFontUnderlineValues value) => Underline[(int)value];
    internal static string ToOpenXmlString(this XlsxFontVerticalTextAlignmentValues value) => VerticalAlignment[(int)value];
    internal static string ToOpenXmlString(this XlsxFontScheme value) => FontScheme[(int)value];
    internal static string ToOpenXmlString(this XlsxPhoneticAlignment value) => PhoneticAlignment[(int)value];
    internal static string ToOpenXmlString(this XlsxPhoneticType value) => PhoneticType[(int)value];
}
