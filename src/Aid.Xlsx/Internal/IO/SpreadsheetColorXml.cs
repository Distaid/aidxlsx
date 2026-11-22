using System.Xml.Linq;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.Internal.IO;

internal static class SpreadsheetColorXml
{
    internal static XlsxColor Read(XElement color)
    {
        if ((string?)color.Attribute("rgb") is { } rgb)
            return XlsxColor.FromColor(ColorStringParser.ParseFromArgb(rgb.AsSpan()));
        if ((uint?)color.Attribute("indexed") is <= 64 and { } indexed)
            return XlsxColor.FromIndex((int)indexed);
        if ((uint?)color.Attribute("theme") is { } theme)
            return XlsxColor.FromTheme((XlsxThemeColor)theme, (double?)color.Attribute("tint") ?? 0);

        return XlsxColor.Automatic;
    }
}
