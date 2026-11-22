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
        switch (themeColor)
        {
            case XlsxThemeColor.Background1:
                return Background1;

            case XlsxThemeColor.Text1:
                return Text1;

            case XlsxThemeColor.Background2:
                return Background2;

            case XlsxThemeColor.Text2:
                return Text2;

            case XlsxThemeColor.Accent1:
                return Accent1;

            case XlsxThemeColor.Accent2:
                return Accent2;

            case XlsxThemeColor.Accent3:
                return Accent3;

            case XlsxThemeColor.Accent4:
                return Accent4;

            case XlsxThemeColor.Accent5:
                return Accent5;

            case XlsxThemeColor.Accent6:
                return Accent6;

            case XlsxThemeColor.Hyperlink:
                return Hyperlink;

            case XlsxThemeColor.FollowedHyperlink:
                return FollowedHyperlink;

            default:
                return null;
        }
    }
}
