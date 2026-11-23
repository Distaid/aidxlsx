#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles;

internal class XlsxTheme : IXlsxTheme
{
    public XlsxColor Background1 { get; set; }
    public XlsxColor Text1 { get; set; }
    public XlsxColor Background2 { get; set; }
    public XlsxColor Text2 { get; set; }
    public XlsxColor Accent1 { get; set; }
    public XlsxColor Accent2 { get; set; }
    public XlsxColor Accent3 { get; set; }
    public XlsxColor Accent4 { get; set; }
    public XlsxColor Accent5 { get; set; }
    public XlsxColor Accent6 { get; set; }
    public XlsxColor Hyperlink { get; set; }
    public XlsxColor FollowedHyperlink { get; set; }

    public XlsxColor ResolveThemeColor(XlsxThemeColor themeColor)
    {
        return themeColor switch
        {
            XlsxThemeColor.Background1 => Background1,
            XlsxThemeColor.Text1 => Text1,
            XlsxThemeColor.Background2 => Background2,
            XlsxThemeColor.Text2 => Text2,
            XlsxThemeColor.Accent1 => Accent1,
            XlsxThemeColor.Accent2 => Accent2,
            XlsxThemeColor.Accent3 => Accent3,
            XlsxThemeColor.Accent4 => Accent4,
            XlsxThemeColor.Accent5 => Accent5,
            XlsxThemeColor.Accent6 => Accent6,
            XlsxThemeColor.Hyperlink => Hyperlink,
            XlsxThemeColor.FollowedHyperlink => FollowedHyperlink,
            _ => null,
        };
    }
}
