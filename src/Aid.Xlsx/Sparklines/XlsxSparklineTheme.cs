#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Sparklines
{
    public static class XlsxSparklineTheme
    {
        #region Public Properties

        public static IXlsxSparklineStyle Default => Dark5;

        #region Dark

        public static IXlsxSparklineStyle Dark1 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.499984740745262),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.249977111117893),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.249977111117893)
        };

        public static IXlsxSparklineStyle Dark2 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.34998626667073579),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.249977111117893),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Dark3 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FF323232"),
            NegativeColor = XlsxColor.FromHtml("FFD00000"),
            MarkersColor = XlsxColor.FromHtml("FFD00000"),
            HighMarkerColor = XlsxColor.FromHtml("FFD00000"),
            LowMarkerColor = XlsxColor.FromHtml("FFD00000"),
            FirstMarkerColor = XlsxColor.FromHtml("FFD00000"),
            LastMarkerColor = XlsxColor.FromHtml("FFD00000")
        };

        public static IXlsxSparklineStyle Dark4 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FF000000"),
            NegativeColor = XlsxColor.FromHtml("FF0070C0"),
            MarkersColor = XlsxColor.FromHtml("FF0070C0"),
            HighMarkerColor = XlsxColor.FromHtml("FF0070C0"),
            LowMarkerColor = XlsxColor.FromHtml("FF0070C0"),
            FirstMarkerColor = XlsxColor.FromHtml("FF0070C0"),
            LastMarkerColor = XlsxColor.FromHtml("FF0070C0")
        };

        public static IXlsxSparklineStyle Dark5 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FF376092"),
            NegativeColor = XlsxColor.FromHtml("FFD00000"),
            MarkersColor = XlsxColor.FromHtml("FFD00000"),
            HighMarkerColor = XlsxColor.FromHtml("FFD00000"),
            LowMarkerColor = XlsxColor.FromHtml("FFD00000"),
            FirstMarkerColor = XlsxColor.FromHtml("FFD00000"),
            LastMarkerColor = XlsxColor.FromHtml("FFD00000")
        };

        public static IXlsxSparklineStyle Dark6 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FF0070C0"),
            NegativeColor = XlsxColor.Black,
            MarkersColor = XlsxColor.Black,
            HighMarkerColor = XlsxColor.Black,
            LowMarkerColor = XlsxColor.Black,
            FirstMarkerColor = XlsxColor.Black,
            LastMarkerColor = XlsxColor.Black
        };

        #endregion Dark

        #region Colorful

        public static IXlsxSparklineStyle Colorful1 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FF5F5F5F"),
            NegativeColor = XlsxColor.FromHtml("FFFFB620"),
            MarkersColor = XlsxColor.FromHtml("FFD70077"),
            HighMarkerColor = XlsxColor.FromHtml("FF56BE79"),
            LowMarkerColor = XlsxColor.FromHtml("FFFF5055"),
            FirstMarkerColor = XlsxColor.FromHtml("FF5687C2"),
            LastMarkerColor = XlsxColor.FromHtml("FF359CEB")
        };

        public static IXlsxSparklineStyle Colorful2 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FF5687C2"),
            NegativeColor = XlsxColor.FromHtml("FFFFB620"),
            MarkersColor = XlsxColor.FromHtml("FFD70077"),
            HighMarkerColor = XlsxColor.FromHtml("FF56BE79"),
            LowMarkerColor = XlsxColor.FromHtml("FFFF5055"),
            FirstMarkerColor = XlsxColor.FromHtml("FF777777"),
            LastMarkerColor = XlsxColor.FromHtml("FF359CEB")
        };

        public static IXlsxSparklineStyle Colorful3 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FFC6EFCE"),
            NegativeColor = XlsxColor.FromHtml("FFFFC7CE"),
            MarkersColor = XlsxColor.FromHtml("FF8CADD6"),
            HighMarkerColor = XlsxColor.FromHtml("FF60D276"),
            LowMarkerColor = XlsxColor.FromHtml("FFFF5367"),
            FirstMarkerColor = XlsxColor.FromHtml("FFFFDC47"),
            LastMarkerColor = XlsxColor.FromHtml("FFFFEB9C")
        };

        public static IXlsxSparklineStyle Colorful4 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromHtml("FF00B050"),
            NegativeColor = XlsxColor.FromHtml("FFFF0000"),
            MarkersColor = XlsxColor.FromHtml("FF0070C0"),
            HighMarkerColor = XlsxColor.FromHtml("FF00B050"),
            LowMarkerColor = XlsxColor.FromHtml("FFFF0000"),
            FirstMarkerColor = XlsxColor.FromHtml("FFFFC000"),
            LastMarkerColor = XlsxColor.FromHtml("FFFFC000")
        };

        public static IXlsxSparklineStyle Colorful5 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Text2),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2)
        };

        public static IXlsxSparklineStyle Colorful6 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Text1),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2)
        };

        #endregion Colorful

        #region Accent

        public static IXlsxSparklineStyle Accent1 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent2 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent3 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent4 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent5 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent6 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893)
        };

        #endregion Accent

        #region Accent Darker 25%

        public static IXlsxSparklineStyle Accent1Darker25 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent2Darker25 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent3Darker25 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent4Darker25 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent5Darker25 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent6Darker25 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893)
        };

        #endregion Accent Darker 25%

        #region Accent Darker 50%

        public static IXlsxSparklineStyle Accent1Darker50 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.499984740745262),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.499984740745262),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, 0.39997558519241921),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, 0.39997558519241921)
        };

        public static IXlsxSparklineStyle Accent2Darker50 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.499984740745262),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.499984740745262),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.39997558519241921),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.39997558519241921)
        };

        public static IXlsxSparklineStyle Accent3Darker50 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.499984740745262),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.499984740745262),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, 0.39997558519241921),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, 0.39997558519241921)
        };

        public static IXlsxSparklineStyle Accent4Darker50 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.499984740745262),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.499984740745262),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, 0.39997558519241921),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, 0.39997558519241921)
        };

        public static IXlsxSparklineStyle Accent5Darker50 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.499984740745262),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.499984740745262),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, 0.39997558519241921),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, 0.39997558519241921)
        };

        public static IXlsxSparklineStyle Accent6Darker50 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.499984740745262),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.499984740745262),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, 0.39997558519241921),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, 0.39997558519241921)
        };

        #endregion Accent Darker 50%

        #region Accent Lighter 40%

        public static IXlsxSparklineStyle Accent1Lighter40 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, 0.39997558519241921),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.499984740745262),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, 0.79998168889431442),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.499984740745262),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.499984740745262),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent1, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent2Lighter40 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.39997558519241921),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.499984740745262),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.79998168889431442),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.499984740745262),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.499984740745262),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent2, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent3Lighter40 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, 0.39997558519241921),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.499984740745262),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, 0.79998168889431442),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.499984740745262),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.499984740745262),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent3, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent4Lighter40 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, 0.39997558519241921),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.499984740745262),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, 0.79998168889431442),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.499984740745262),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.499984740745262),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent4, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent5Lighter40 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, 0.39997558519241921),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.499984740745262),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, 0.79998168889431442),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.499984740745262),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.499984740745262),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent5, -0.249977111117893)
        };

        public static IXlsxSparklineStyle Accent6Lighter40 => new XlsxSparklineStyle
        {
            SeriesColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, 0.39997558519241921),
            NegativeColor = XlsxColor.FromTheme(XlsxThemeColor.Background1, -0.499984740745262),
            MarkersColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, 0.79998168889431442),
            HighMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.499984740745262),
            LowMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.499984740745262),
            FirstMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893),
            LastMarkerColor = XlsxColor.FromTheme(XlsxThemeColor.Accent6, -0.249977111117893)
        };

        #endregion Accent Lighter 40%

        #endregion Public Properties
    }
}
