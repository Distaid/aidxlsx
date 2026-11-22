using System.Xml.Linq;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.SpreadsheetML.Styles;

internal static class ThemeXml
{
    private static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Read(Stream stream, IXlsxTheme theme)
    {
        var scheme = GetScheme(XDocumentExtensions.Load(stream));
        if (scheme is null)
            return;

        theme.Text1 = ReadColor(scheme, "dk1") ?? theme.Text1;
        theme.Background1 = ReadColor(scheme, "lt1") ?? theme.Background1;
        theme.Text2 = ReadColor(scheme, "dk2") ?? theme.Text2;
        theme.Background2 = ReadColor(scheme, "lt2") ?? theme.Background2;
        theme.Accent1 = ReadColor(scheme, "accent1") ?? theme.Accent1;
        theme.Accent2 = ReadColor(scheme, "accent2") ?? theme.Accent2;
        theme.Accent3 = ReadColor(scheme, "accent3") ?? theme.Accent3;
        theme.Accent4 = ReadColor(scheme, "accent4") ?? theme.Accent4;
        theme.Accent5 = ReadColor(scheme, "accent5") ?? theme.Accent5;
        theme.Accent6 = ReadColor(scheme, "accent6") ?? theme.Accent6;
        theme.Hyperlink = ReadColor(scheme, "hlink") ?? theme.Hyperlink;
        theme.FollowedHyperlink = ReadColor(scheme, "folHlink") ?? theme.FollowedHyperlink;
    }

    internal static void WriteDefault(Stream stream, IXlsxTheme theme)
    {
        using var template = typeof(ThemeXml).Assembly.GetManifestResourceStream("Aid.Xlsx.DefaultTheme.xml")
            ?? throw new InvalidOperationException("The default workbook theme is missing.");
        var xml = XDocument.Load(template);
        var scheme = GetScheme(xml)!;

        SetColor("dk1", theme.Text1, "sysClr", "lastClr");
        SetColor("lt1", theme.Background1, "sysClr", "lastClr");
        SetColor("dk2", theme.Text2);
        SetColor("lt2", theme.Background2);
        SetColor("accent1", theme.Accent1);
        SetColor("accent2", theme.Accent2);
        SetColor("accent3", theme.Accent3);
        SetColor("accent4", theme.Accent4);
        SetColor("accent5", theme.Accent5);
        SetColor("accent6", theme.Accent6);
        SetColor("hlink", theme.Hyperlink);
        SetColor("folHlink", theme.FollowedHyperlink);
        xml.Save(stream, System.Xml.Linq.SaveOptions.DisableFormatting);

        void SetColor(string name, XlsxColor color, string element = "srgbClr", string attribute = "val")
        {
            scheme.Element(Drawing + name)!.Element(Drawing + element)!
                .SetAttributeValue(attribute, color.Color.ToHex()[2..]);
        }
    }

    private static XElement? GetScheme(XDocument? xml) =>
        xml?.Root?.Element(Drawing + "themeElements")?.Element(Drawing + "clrScheme");

    private static XlsxColor? ReadColor(XElement scheme, string name)
    {
        var slot = scheme.Element(Drawing + name);
        var rgb = (string?)slot?.Element(Drawing + "srgbClr")?.Attribute("val")
            ?? (string?)slot?.Element(Drawing + "sysClr")?.Attribute("lastClr");
        return string.IsNullOrEmpty(rgb) ? null : XlsxColor.FromHexRgb(rgb);
    }
}
