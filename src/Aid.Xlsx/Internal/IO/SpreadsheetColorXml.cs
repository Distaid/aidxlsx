using System.Xml.Linq;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.Internal.IO;

internal static class SpreadsheetColorXml
{
    internal static XElement Write(XName name, XlsxColor color)
    {
        var element = new XElement(name);
        switch (color.ColorType)
        {
            case XlsxColorType.Color:
                element.SetAttributeValue("rgb", color.Color.ToHex());
                break;
            case XlsxColorType.Indexed:
                element.SetAttributeValue("indexed", color.Indexed);
                break;
            case XlsxColorType.Theme:
                element.SetAttributeValue("theme", (int)color.ThemeColor);
                if (color.ThemeTint != 0)
                    element.SetAttributeValue("tint", color.ThemeTint);
                break;
        }

        return element;
    }

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
