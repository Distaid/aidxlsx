#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles;

public interface IXlsxTheme
{
    XlsxColor Background1 { get; set; }
    XlsxColor Text1 { get; set; }
    XlsxColor Background2 { get; set; }
    XlsxColor Text2 { get; set; }
    XlsxColor Accent1 { get; set; }
    XlsxColor Accent2 { get; set; }
    XlsxColor Accent3 { get; set; }
    XlsxColor Accent4 { get; set; }
    XlsxColor Accent5 { get; set; }
    XlsxColor Accent6 { get; set; }
    XlsxColor Hyperlink { get; set; }
    XlsxColor FollowedHyperlink { get; set; }

    XlsxColor ResolveThemeColor(XlsxThemeColor themeColor);
}
