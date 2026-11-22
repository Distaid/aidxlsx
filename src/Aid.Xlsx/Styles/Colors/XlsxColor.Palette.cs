using System.Drawing;
using Aid.Xlsx.Caching;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.Styles.Colors
{
    public partial class XlsxColor
    {
        private static readonly XlsxColorRepository Repository = new XlsxColorRepository(key => new XlsxColor(key));

        /// <summary>
        /// System foreground color. Depends on system settings, but there is no platform-agnostic way to resolve it.
        /// </summary>
        private static readonly XlsxColor SystemForeground = FromRgb(0x000000);

        /// <summary>
        /// System background color. Depends on system settings, but there is no platform-agnostic way to resolve it.
        /// </summary>
        private static readonly XlsxColor SystemBackground = FromRgb(0xFFFFFF);

        /// <summary>
        /// VML palette entries from MS-OI29500. Excel uses Windows system color scheme to determine the actual colors of a palette
        /// entry, but we have no way to get them. Win10 doesn't even have a tool, use Classic Color Panel. We will use the default
        /// values that are default on Windows.
        /// </summary>
        private static readonly Lazy<Dictionary<String, XlsxColor>> VmlPaletteColors = new(() => new Dictionary<String, XlsxColor>(StringComparer.OrdinalIgnoreCase)
        {
            { "ButtonFace", FromRgb(0xF0F0F0) },
            { "WindowText", SystemForeground },
            { "Menu", FromRgb(0xF0F0F0) },
            { "Highlight", FromRgb(0x0078D7) },
            { "HighlightText", FromRgb(0xFFFFFF) },
            { "CaptionText", FromRgb(0x000000) },
            { "ActiveCaption", FromRgb(0x99B4D1) },
            { "ButtonHighlight", FromRgb(0xFFFFFF) },
            { "ButtonShadow", FromRgb(0xA0A0A0) },
            { "ButtonText", FromRgb(0x000000) },
            { "GrayText", FromRgb(0x6D6D6D) },
            { "InactiveCaption", FromRgb(0xBFCDDB) },
            { "InactiveCaptionText", FromRgb(0x000000) },
            { "InfoBackground", FromRgb(0xFFFFE1) },
            { "InfoText", FromRgb(0x000000) },
            { "MenuText", FromRgb(0x000000) },
            { "Scrollbar", FromRgb(0xC8C8C8) },
            { "Window", SystemBackground },
            { "WindowFrame", FromRgb(0x646464) },
            { "ThreeDLightShadow", FromRgb(0x000000) },
            { "ThreeDDarkShadow", FromRgb(0x696969) },
            { "ActiveBorder", FromRgb(0xB4B4B4) },
            { "InactiveBorder", FromRgb(0xF4F7FC) },
            { "Background", FromRgb(0x000000) },
            { "AppWorkspace", FromRgb(0xABABAB) },
            { "ThreeDFace", FromRgb(0xF0F0F0) },
            { "ThreeDShadow", FromRgb(0xA0A0A0) },
            { "ThreeDHighlight", FromRgb(0xFFFFFF) }
        });

        /// <summary>
        /// Named colors for VML ST_ColorType from ISO-29500:4.
        /// </summary>
        private static readonly Lazy<Dictionary<string, XlsxColor>> VmlNamedColors = new(() => new Dictionary<string, XlsxColor>(StringComparer.OrdinalIgnoreCase)
        {
            { "Black", Black },
            { "Silver", Silver },
            { "Gray", Gray },
            { "White", White },
            { "Maroon", Maroon },
            { "Red", Red },
            { "Purple", Purple },
            { "Fuchsia", Fuchsia },
            { "Green", Green  },
            { "Lime", Lime },
            { "Olive", Olive },
            { "Yellow", Yellow },
            { "Navy", Navy },
            { "Blue", Blue },
            { "Teal", Teal },
            { "Aqua", Aqua },
        });

        /// <summary>
        /// Default indexed colors from ISO-29500:18.8.27. If a workbook doesn't have a specific colors, these are used.
        /// </summary>
        private static readonly Lazy<Dictionary<Int32, XlsxColor>> IndexedColorsLazy = new(() => new Dictionary<int, XlsxColor>
        {
            {  0, FromArgb(0xFF000000) },
            {  1, FromArgb(0xFFFFFFFF) },
            {  2, FromArgb(0xFFFF0000) },
            {  3, FromArgb(0xFF00FF00) },
            {  4, FromArgb(0xFF0000FF) },
            {  5, FromArgb(0xFFFFFF00) },
            {  6, FromArgb(0xFFFF00FF) },
            {  7, FromArgb(0xFF00FFFF) },
            {  8, FromArgb(0xFF000000) },
            {  9, FromArgb(0xFFFFFFFF) },
            { 10, FromArgb(0xFFFF0000) },
            { 11, FromArgb(0xFF00FF00) },
            { 12, FromArgb(0xFF0000FF) },
            { 13, FromArgb(0xFFFFFF00) },
            { 14, FromArgb(0xFFFF00FF) },
            { 15, FromArgb(0xFF00FFFF) },
            { 16, FromArgb(0xFF800000) },
            { 17, FromArgb(0xFF008000) },
            { 18, FromArgb(0xFF000080) },
            { 19, FromArgb(0xFF808000) },
            { 20, FromArgb(0xFF800080) },
            { 21, FromArgb(0xFF008080) },
            { 22, FromArgb(0xFFC0C0C0) },
            { 23, FromArgb(0xFF808080) },
            { 24, FromArgb(0xFF9999FF) },
            { 25, FromArgb(0xFF993366) },
            { 26, FromArgb(0xFFFFFFCC) },
            { 27, FromArgb(0xFFCCFFFF) },
            { 28, FromArgb(0xFF660066) },
            { 29, FromArgb(0xFFFF8080) },
            { 30, FromArgb(0xFF0066CC) },
            { 31, FromArgb(0xFFCCCCFF) },
            { 32, FromArgb(0xFF000080) },
            { 33, FromArgb(0xFFFF00FF) },
            { 34, FromArgb(0xFFFFFF00) },
            { 35, FromArgb(0xFF00FFFF) },
            { 36, FromArgb(0xFF800080) },
            { 37, FromArgb(0xFF800000) },
            { 38, FromArgb(0xFF008080) },
            { 39, FromArgb(0xFF0000FF) },
            { 40, FromArgb(0xFF00CCFF) },
            { 41, FromArgb(0xFFCCFFFF) },
            { 42, FromArgb(0xFFCCFFCC) },
            { 43, FromArgb(0xFFFFFF99) },
            { 44, FromArgb(0xFF99CCFF) },
            { 45, FromArgb(0xFFFF99CC) },
            { 46, FromArgb(0xFFCC99FF) },
            { 47, FromArgb(0xFFFFCC99) },
            { 48, FromArgb(0xFF3366FF) },
            { 49, FromArgb(0xFF33CCCC) },
            { 50, FromArgb(0xFF99CC00) },
            { 51, FromArgb(0xFFFFCC00) },
            { 52, FromArgb(0xFFFF9900) },
            { 53, FromArgb(0xFFFF6600) },
            { 54, FromArgb(0xFF666699) },
            { 55, FromArgb(0xFF969696) },
            { 56, FromArgb(0xFF003366) },
            { 57, FromArgb(0xFF339966) },
            { 58, FromArgb(0xFF003300) },
            { 59, FromArgb(0xFF333300) },
            { 60, FromArgb(0xFF993300) },
            { 61, FromArgb(0xFF993366) },
            { 62, FromArgb(0xFF333399) },
            { 63, FromArgb(0xFF333333) },
            { 64, SystemForeground },
            { 65, SystemBackground },
        });

        /// <summary>
        /// Default indexed colors per ISO-298500 indexedColors. All values are either RGB or named.
        /// </summary>
        public static Dictionary<Int32, XlsxColor> IndexedColors => IndexedColorsLazy.Value;

        internal static XlsxColor FromKey(ref XlsxColorKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        public static XlsxColor FromColor(Color color)
        {
            var key = new XlsxColorKey
            {
                ColorType = XlsxColorType.Color,
                Color = color
            };
            return FromKey(ref key);
        }

        public static XlsxColor FromArgb(Int32 argb)
        {
            return FromColor(Color.FromArgb(argb));
        }

        internal static XlsxColor FromArgb(UInt32 argb)
        {
            return FromColor(Color.FromArgb(unchecked((int)argb)));
        }

        public static XlsxColor FromArgb(Int32 r, Int32 g, Int32 b)
        {
            return FromColor(Color.FromArgb(r, g, b));
        }

        public static XlsxColor FromArgb(Int32 a, Int32 r, Int32 g, Int32 b)
        {
            return FromColor(Color.FromArgb(a, r, g, b));
        }

        /// <summary>
        /// Create a color from RBG hexa number. Alpha will be always set to full opacity.
        /// </summary>
        internal static XlsxColor FromRgb(Int32 rgb)
        {
            unchecked
            {
                return FromColor(Color.FromArgb(rgb | (int)0xFF000000));
            }
        }

        /// <summary>
        /// Parse a color in a RRGGBB hexadecimal format. It is used
        /// to parse colors of <c>ST_HexColorRGB</c> type from the XML.
        /// </summary>
        /// <param name="hexColorRgb">A 6 character long hexadecimal number.</param>
        /// <returns>Parsed color with full opacity.</returns>
        internal static XlsxColor FromHexRgb(String hexColorRgb)
        {
            return FromColor(ColorStringParser.ParseFromRgb6(hexColorRgb));
        }

        public static XlsxColor FromName(String name)
        {
            return FromColor(Color.FromName(name));
        }

        public static XlsxColor FromHtml(String htmlColor)
        {
            return FromColor(ColorStringParser.ParseFromHtml(htmlColor));
        }

        public static XlsxColor FromIndex(Int32 index)
        {
            var key = new XlsxColorKey
            {
                ColorType = XlsxColorType.Indexed,
                Indexed = index
            };
            return FromKey(ref key);
        }

        public static XlsxColor FromTheme(XlsxThemeColor themeColor)
        {
            var key = new XlsxColorKey
            {
                ColorType = XlsxColorType.Theme,
                ThemeColor = themeColor
            };
            return FromKey(ref key);
        }

        public static XlsxColor FromTheme(XlsxThemeColor themeColor, Double themeTint)
        {
            var key = new XlsxColorKey
            {
                ColorType = XlsxColorType.Theme,
                ThemeColor = themeColor,
                ThemeTint = themeTint
            };
            return FromKey(ref key);
        }

        /// <summary>
        /// Parse VML ST_ColorType type from ECMA-376, Part 4 §20.1.2.3.
        /// </summary>
        internal static XlsxColor FromVmlColor(String vmlColor)
        {
            // Check if VML color is a hexadecimal RGB color.
            var rgbColorText = vmlColor.AsSpan().Trim();
            if (rgbColorText.StartsWith("#".AsSpan()) && ColorStringParser.TryParseRgb6(rgbColorText[1..], out var rgbColor))
            {
                return FromColor(rgbColor);
            }

            // Check if VML color is a VML named color.
            var namedColorText = vmlColor.Trim();
            if (VmlNamedColors.Value.TryGetValue(namedColorText, out var namedColor))
            {
                return namedColor;
            }

            // Check if VML color is a palette color. ST_ColorType palette entry can have an *optional* index to
            // a palette (that is a *system* palette, not indexed colors). The palette index only brings trouble
            // because VML is used pretty much only for unusual features like notes, form controls ect. Strip the
            // palette index if found, it is a legacy feature from times of 256 colors. E.g., "Menu [80]" => "Menu"
            var paletteColorName = vmlColor.Split('[')[0].Trim();
            if (VmlPaletteColors.Value.TryGetValue(paletteColorName, out var paletteColor))
            {
                return paletteColor;
            }

            // Don't crash on unparsable colors, e.g. None or malformed #RED should still be represented, even
            // though we don't actually know how to map them.
            return FromName(vmlColor);
        }

        /// <summary>
        /// Automatic color. The color is selected by the Excel depending on a context of usage.
        /// Example: black for font color, white for background and black for pattern in a pattern
        /// fill.
        /// </summary>
        public static XlsxColor Automatic { get; } = new XlsxColor();

        /// <summary>
        /// An alias for <see cref="Automatic"/> color. In some Excel color pickers (e.g. sheet tab
        /// color picker or fill background color), the GUI displays "No Color" option instead of
        /// "Automatic" button for the automatic color. That is only GUI convention, the stored
        /// value is automatic color.
        /// </summary>
        public static XlsxColor NoColor => Automatic;

        public static XlsxColor AliceBlue { get { return FromColor(Color.AliceBlue); } }

        public static XlsxColor AntiqueWhite { get { return FromColor(Color.AntiqueWhite); } }

        public static XlsxColor Aqua { get { return FromColor(Color.Aqua); } }

        public static XlsxColor Aquamarine { get { return FromColor(Color.Aquamarine); } }

        public static XlsxColor Azure { get { return FromColor(Color.Azure); } }

        public static XlsxColor Beige { get { return FromColor(Color.Beige); } }

        public static XlsxColor Bisque { get { return FromColor(Color.Bisque); } }

        public static XlsxColor Black { get { return FromColor(Color.Black); } }

        public static XlsxColor BlanchedAlmond { get { return FromColor(Color.BlanchedAlmond); } }

        public static XlsxColor Blue { get { return FromColor(Color.Blue); } }

        public static XlsxColor BlueViolet { get { return FromColor(Color.BlueViolet); } }

        public static XlsxColor Brown { get { return FromColor(Color.Brown); } }

        public static XlsxColor BurlyWood { get { return FromColor(Color.BurlyWood); } }

        public static XlsxColor CadetBlue { get { return FromColor(Color.CadetBlue); } }

        public static XlsxColor Chartreuse { get { return FromColor(Color.Chartreuse); } }

        public static XlsxColor Chocolate { get { return FromColor(Color.Chocolate); } }

        public static XlsxColor Coral { get { return FromColor(Color.Coral); } }

        public static XlsxColor CornflowerBlue { get { return FromColor(Color.CornflowerBlue); } }

        public static XlsxColor Cornsilk { get { return FromColor(Color.Cornsilk); } }

        public static XlsxColor Crimson { get { return FromColor(Color.Crimson); } }

        public static XlsxColor Cyan { get { return FromColor(Color.Cyan); } }

        public static XlsxColor DarkBlue { get { return FromColor(Color.DarkBlue); } }

        public static XlsxColor DarkCyan { get { return FromColor(Color.DarkCyan); } }

        public static XlsxColor DarkGoldenrod { get { return FromColor(Color.DarkGoldenrod); } }

        public static XlsxColor DarkGray { get { return FromColor(Color.DarkGray); } }

        public static XlsxColor DarkGreen { get { return FromColor(Color.DarkGreen); } }

        public static XlsxColor DarkKhaki { get { return FromColor(Color.DarkKhaki); } }

        public static XlsxColor DarkMagenta { get { return FromColor(Color.DarkMagenta); } }

        public static XlsxColor DarkOliveGreen { get { return FromColor(Color.DarkOliveGreen); } }

        public static XlsxColor DarkOrange { get { return FromColor(Color.DarkOrange); } }

        public static XlsxColor DarkOrchid { get { return FromColor(Color.DarkOrchid); } }

        public static XlsxColor DarkRed { get { return FromColor(Color.DarkRed); } }

        public static XlsxColor DarkSalmon { get { return FromColor(Color.DarkSalmon); } }

        public static XlsxColor DarkSeaGreen { get { return FromColor(Color.DarkSeaGreen); } }

        public static XlsxColor DarkSlateBlue { get { return FromColor(Color.DarkSlateBlue); } }

        public static XlsxColor DarkSlateGray { get { return FromColor(Color.DarkSlateGray); } }

        public static XlsxColor DarkTurquoise { get { return FromColor(Color.DarkTurquoise); } }

        public static XlsxColor DarkViolet { get { return FromColor(Color.DarkViolet); } }

        public static XlsxColor DeepPink { get { return FromColor(Color.DeepPink); } }

        public static XlsxColor DeepSkyBlue { get { return FromColor(Color.DeepSkyBlue); } }

        public static XlsxColor DimGray { get { return FromColor(Color.DimGray); } }

        public static XlsxColor DodgerBlue { get { return FromColor(Color.DodgerBlue); } }

        public static XlsxColor Firebrick { get { return FromColor(Color.Firebrick); } }

        public static XlsxColor FloralWhite { get { return FromColor(Color.FloralWhite); } }

        public static XlsxColor ForestGreen { get { return FromColor(Color.ForestGreen); } }

        public static XlsxColor Fuchsia { get { return FromColor(Color.Fuchsia); } }

        public static XlsxColor Gainsboro { get { return FromColor(Color.Gainsboro); } }

        public static XlsxColor GhostWhite { get { return FromColor(Color.GhostWhite); } }

        public static XlsxColor Gold { get { return FromColor(Color.Gold); } }

        public static XlsxColor Goldenrod { get { return FromColor(Color.Goldenrod); } }

        public static XlsxColor Gray { get { return FromColor(Color.Gray); } }

        public static XlsxColor Green { get { return FromColor(Color.Green); } }

        public static XlsxColor GreenYellow { get { return FromColor(Color.GreenYellow); } }

        public static XlsxColor Honeydew { get { return FromColor(Color.Honeydew); } }

        public static XlsxColor HotPink { get { return FromColor(Color.HotPink); } }

        public static XlsxColor IndianRed { get { return FromColor(Color.IndianRed); } }

        public static XlsxColor Indigo { get { return FromColor(Color.Indigo); } }

        public static XlsxColor Ivory { get { return FromColor(Color.Ivory); } }

        public static XlsxColor Khaki { get { return FromColor(Color.Khaki); } }

        public static XlsxColor Lavender { get { return FromColor(Color.Lavender); } }

        public static XlsxColor LavenderBlush { get { return FromColor(Color.LavenderBlush); } }

        public static XlsxColor LawnGreen { get { return FromColor(Color.LawnGreen); } }

        public static XlsxColor LemonChiffon { get { return FromColor(Color.LemonChiffon); } }

        public static XlsxColor LightBlue { get { return FromColor(Color.LightBlue); } }

        public static XlsxColor LightCoral { get { return FromColor(Color.LightCoral); } }

        public static XlsxColor LightCyan { get { return FromColor(Color.LightCyan); } }

        public static XlsxColor LightGoldenrodYellow { get { return FromColor(Color.LightGoldenrodYellow); } }

        public static XlsxColor LightGray { get { return FromColor(Color.LightGray); } }

        public static XlsxColor LightGreen { get { return FromColor(Color.LightGreen); } }

        public static XlsxColor LightPink { get { return FromColor(Color.LightPink); } }

        public static XlsxColor LightSalmon { get { return FromColor(Color.LightSalmon); } }

        public static XlsxColor LightSeaGreen { get { return FromColor(Color.LightSeaGreen); } }

        public static XlsxColor LightSkyBlue { get { return FromColor(Color.LightSkyBlue); } }

        public static XlsxColor LightSlateGray { get { return FromColor(Color.LightSlateGray); } }

        public static XlsxColor LightSteelBlue { get { return FromColor(Color.LightSteelBlue); } }

        public static XlsxColor LightYellow { get { return FromColor(Color.LightYellow); } }

        public static XlsxColor Lime { get { return FromColor(Color.Lime); } }

        public static XlsxColor LimeGreen { get { return FromColor(Color.LimeGreen); } }

        public static XlsxColor Linen { get { return FromColor(Color.Linen); } }

        public static XlsxColor Magenta { get { return FromColor(Color.Magenta); } }

        public static XlsxColor Maroon { get { return FromColor(Color.Maroon); } }

        public static XlsxColor MediumAquamarine { get { return FromColor(Color.MediumAquamarine); } }

        public static XlsxColor MediumBlue { get { return FromColor(Color.MediumBlue); } }

        public static XlsxColor MediumOrchid { get { return FromColor(Color.MediumOrchid); } }

        public static XlsxColor MediumPurple { get { return FromColor(Color.MediumPurple); } }

        public static XlsxColor MediumSeaGreen { get { return FromColor(Color.MediumSeaGreen); } }

        public static XlsxColor MediumSlateBlue { get { return FromColor(Color.MediumSlateBlue); } }

        public static XlsxColor MediumSpringGreen { get { return FromColor(Color.MediumSpringGreen); } }

        public static XlsxColor MediumTurquoise { get { return FromColor(Color.MediumTurquoise); } }

        public static XlsxColor MediumVioletRed { get { return FromColor(Color.MediumVioletRed); } }

        public static XlsxColor MidnightBlue { get { return FromColor(Color.MidnightBlue); } }

        public static XlsxColor MintCream { get { return FromColor(Color.MintCream); } }

        public static XlsxColor MistyRose { get { return FromColor(Color.MistyRose); } }

        public static XlsxColor Moccasin { get { return FromColor(Color.Moccasin); } }

        public static XlsxColor NavajoWhite { get { return FromColor(Color.NavajoWhite); } }

        public static XlsxColor Navy { get { return FromColor(Color.Navy); } }

        public static XlsxColor OldLace { get { return FromColor(Color.OldLace); } }

        public static XlsxColor Olive { get { return FromColor(Color.Olive); } }

        public static XlsxColor OliveDrab { get { return FromColor(Color.OliveDrab); } }

        public static XlsxColor Orange { get { return FromColor(Color.Orange); } }

        public static XlsxColor OrangeRed { get { return FromColor(Color.OrangeRed); } }

        public static XlsxColor Orchid { get { return FromColor(Color.Orchid); } }

        public static XlsxColor PaleGoldenrod { get { return FromColor(Color.PaleGoldenrod); } }

        public static XlsxColor PaleGreen { get { return FromColor(Color.PaleGreen); } }

        public static XlsxColor PaleTurquoise { get { return FromColor(Color.PaleTurquoise); } }

        public static XlsxColor PaleVioletRed { get { return FromColor(Color.PaleVioletRed); } }

        public static XlsxColor PapayaWhip { get { return FromColor(Color.PapayaWhip); } }

        public static XlsxColor PeachPuff { get { return FromColor(Color.PeachPuff); } }

        public static XlsxColor Peru { get { return FromColor(Color.Peru); } }

        public static XlsxColor Pink { get { return FromColor(Color.Pink); } }

        public static XlsxColor Plum { get { return FromColor(Color.Plum); } }

        public static XlsxColor PowderBlue { get { return FromColor(Color.PowderBlue); } }

        public static XlsxColor Purple { get { return FromColor(Color.Purple); } }

        public static XlsxColor Red { get { return FromColor(Color.Red); } }

        public static XlsxColor RosyBrown { get { return FromColor(Color.RosyBrown); } }

        public static XlsxColor RoyalBlue { get { return FromColor(Color.RoyalBlue); } }

        public static XlsxColor SaddleBrown { get { return FromColor(Color.SaddleBrown); } }

        public static XlsxColor Salmon { get { return FromColor(Color.Salmon); } }

        public static XlsxColor SandyBrown { get { return FromColor(Color.SandyBrown); } }

        public static XlsxColor SeaGreen { get { return FromColor(Color.SeaGreen); } }

        public static XlsxColor SeaShell { get { return FromColor(Color.SeaShell); } }

        public static XlsxColor Sienna { get { return FromColor(Color.Sienna); } }

        public static XlsxColor Silver { get { return FromColor(Color.Silver); } }

        public static XlsxColor SkyBlue { get { return FromColor(Color.SkyBlue); } }

        public static XlsxColor SlateBlue { get { return FromColor(Color.SlateBlue); } }

        public static XlsxColor SlateGray { get { return FromColor(Color.SlateGray); } }

        public static XlsxColor Snow { get { return FromColor(Color.Snow); } }

        public static XlsxColor SpringGreen { get { return FromColor(Color.SpringGreen); } }

        public static XlsxColor SteelBlue { get { return FromColor(Color.SteelBlue); } }

        public static XlsxColor Tan { get { return FromColor(Color.Tan); } }

        public static XlsxColor Teal { get { return FromColor(Color.Teal); } }

        public static XlsxColor Thistle { get { return FromColor(Color.Thistle); } }

        public static XlsxColor Tomato { get { return FromColor(Color.Tomato); } }

        public static XlsxColor Turquoise { get { return FromColor(Color.Turquoise); } }

        public static XlsxColor Violet { get { return FromColor(Color.Violet); } }

        public static XlsxColor Wheat { get { return FromColor(Color.Wheat); } }

        public static XlsxColor White { get { return FromColor(Color.White); } }

        public static XlsxColor WhiteSmoke { get { return FromColor(Color.WhiteSmoke); } }

        public static XlsxColor Yellow { get { return FromColor(Color.Yellow); } }

        public static XlsxColor YellowGreen { get { return FromColor(Color.YellowGreen); } }

        public static XlsxColor AirForceBlue { get { return FromHtml("#FF5D8AA8"); } }

        public static XlsxColor Alizarin { get { return FromHtml("#FFE32636"); } }

        public static XlsxColor Almond { get { return FromHtml("#FFEFDECD"); } }

        public static XlsxColor Amaranth { get { return FromHtml("#FFE52B50"); } }

        public static XlsxColor Amber { get { return FromHtml("#FFFFBF00"); } }

        public static XlsxColor AmberSaeEce { get { return FromHtml("#FFFF7E00"); } }

        public static XlsxColor AmericanRose { get { return FromHtml("#FFFF033E"); } }

        public static XlsxColor Amethyst { get { return FromHtml("#FF9966CC"); } }

        public static XlsxColor AntiFlashWhite { get { return FromHtml("#FFF2F3F4"); } }

        public static XlsxColor AntiqueBrass { get { return FromHtml("#FFCD9575"); } }

        public static XlsxColor AntiqueFuchsia { get { return FromHtml("#FF915C83"); } }

        public static XlsxColor AppleGreen { get { return FromHtml("#FF8DB600"); } }

        public static XlsxColor Apricot { get { return FromHtml("#FFFBCEB1"); } }

        public static XlsxColor Aquamarine1 { get { return FromHtml("#FF7FFFD0"); } }

        public static XlsxColor ArmyGreen { get { return FromHtml("#FF4B5320"); } }

        public static XlsxColor Arsenic { get { return FromHtml("#FF3B444B"); } }

        public static XlsxColor ArylideYellow { get { return FromHtml("#FFE9D66B"); } }

        public static XlsxColor AshGrey { get { return FromHtml("#FFB2BEB5"); } }

        public static XlsxColor Asparagus { get { return FromHtml("#FF87A96B"); } }

        public static XlsxColor AtomicTangerine { get { return FromHtml("#FFFF9966"); } }

        public static XlsxColor Auburn { get { return FromHtml("#FF6D351A"); } }

        public static XlsxColor Aureolin { get { return FromHtml("#FFFDEE00"); } }

        public static XlsxColor Aurometalsaurus { get { return FromHtml("#FF6E7F80"); } }

        public static XlsxColor Awesome { get { return FromHtml("#FFFF2052"); } }

        public static XlsxColor AzureColorWheel { get { return FromHtml("#FF007FFF"); } }

        public static XlsxColor BabyBlue { get { return FromHtml("#FF89CFF0"); } }

        public static XlsxColor BabyBlueEyes { get { return FromHtml("#FFA1CAF1"); } }

        public static XlsxColor BabyPink { get { return FromHtml("#FFF4C2C2"); } }

        public static XlsxColor BallBlue { get { return FromHtml("#FF21ABCD"); } }

        public static XlsxColor BananaMania { get { return FromHtml("#FFFAE7B5"); } }

        public static XlsxColor BattleshipGrey { get { return FromHtml("#FF848482"); } }

        public static XlsxColor Bazaar { get { return FromHtml("#FF98777B"); } }

        public static XlsxColor BeauBlue { get { return FromHtml("#FFBCD4E6"); } }

        public static XlsxColor Beaver { get { return FromHtml("#FF9F8170"); } }

        public static XlsxColor Bistre { get { return FromHtml("#FF3D2B1F"); } }

        public static XlsxColor Bittersweet { get { return FromHtml("#FFFE6F5E"); } }

        public static XlsxColor BleuDeFrance { get { return FromHtml("#FF318CE7"); } }

        public static XlsxColor BlizzardBlue { get { return FromHtml("#FFACE5EE"); } }

        public static XlsxColor Blond { get { return FromHtml("#FFFAF0BE"); } }

        public static XlsxColor BlueBell { get { return FromHtml("#FFA2A2D0"); } }

        public static XlsxColor BlueGray { get { return FromHtml("#FF6699CC"); } }

        public static XlsxColor BlueGreen { get { return FromHtml("#FF00DDDD"); } }

        public static XlsxColor BluePigment { get { return FromHtml("#FF333399"); } }

        public static XlsxColor BlueRyb { get { return FromHtml("#FF0247FE"); } }

        public static XlsxColor Blush { get { return FromHtml("#FFDE5D83"); } }

        public static XlsxColor Bole { get { return FromHtml("#FF79443B"); } }

        public static XlsxColor BondiBlue { get { return FromHtml("#FF0095B6"); } }

        public static XlsxColor BostonUniversityRed { get { return FromHtml("#FFCC0000"); } }

        public static XlsxColor BrandeisBlue { get { return FromHtml("#FF0070FF"); } }

        public static XlsxColor Brass { get { return FromHtml("#FFB5A642"); } }

        public static XlsxColor BrickRed { get { return FromHtml("#FFCB4154"); } }

        public static XlsxColor BrightCerulean { get { return FromHtml("#FF1DACD6"); } }

        public static XlsxColor BrightGreen { get { return FromHtml("#FF66FF00"); } }

        public static XlsxColor BrightLavender { get { return FromHtml("#FFBF94E4"); } }

        public static XlsxColor BrightMaroon { get { return FromHtml("#FFC32148"); } }

        public static XlsxColor BrightPink { get { return FromHtml("#FFFF007F"); } }

        public static XlsxColor BrightTurquoise { get { return FromHtml("#FF08E8DE"); } }

        public static XlsxColor BrightUbe { get { return FromHtml("#FFD19FE8"); } }

        public static XlsxColor BrilliantLavender { get { return FromHtml("#FFF4BBFF"); } }

        public static XlsxColor BrilliantRose { get { return FromHtml("#FFFF55A3"); } }

        public static XlsxColor BrinkPink { get { return FromHtml("#FFFB607F"); } }

        public static XlsxColor BritishRacingGreen { get { return FromHtml("#FF004225"); } }

        public static XlsxColor Bronze { get { return FromHtml("#FFCD7F32"); } }

        public static XlsxColor BrownTraditional { get { return FromHtml("#FF964B00"); } }

        public static XlsxColor BubbleGum { get { return FromHtml("#FFFFC1CC"); } }

        public static XlsxColor Bubbles { get { return FromHtml("#FFE7FEFF"); } }

        public static XlsxColor Buff { get { return FromHtml("#FFF0DC82"); } }

        public static XlsxColor BulgarianRose { get { return FromHtml("#FF480607"); } }

        public static XlsxColor Burgundy { get { return FromHtml("#FF800020"); } }

        public static XlsxColor BurntOrange { get { return FromHtml("#FFCC5500"); } }

        public static XlsxColor BurntSienna { get { return FromHtml("#FFE97451"); } }

        public static XlsxColor BurntUmber { get { return FromHtml("#FF8A3324"); } }

        public static XlsxColor Byzantine { get { return FromHtml("#FFBD33A4"); } }

        public static XlsxColor Byzantium { get { return FromHtml("#FF702963"); } }

        public static XlsxColor Cadet { get { return FromHtml("#FF536872"); } }

        public static XlsxColor CadetGrey { get { return FromHtml("#FF91A3B0"); } }

        public static XlsxColor CadmiumGreen { get { return FromHtml("#FF006B3C"); } }

        public static XlsxColor CadmiumOrange { get { return FromHtml("#FFED872D"); } }

        public static XlsxColor CadmiumRed { get { return FromHtml("#FFE30022"); } }

        public static XlsxColor CadmiumYellow { get { return FromHtml("#FFFFF600"); } }

        public static XlsxColor CalPolyPomonaGreen { get { return FromHtml("#FF1E4D2B"); } }

        public static XlsxColor CambridgeBlue { get { return FromHtml("#FFA3C1AD"); } }

        public static XlsxColor Camel { get { return FromHtml("#FFC19A6B"); } }

        public static XlsxColor CamouflageGreen { get { return FromHtml("#FF78866B"); } }

        public static XlsxColor CanaryYellow { get { return FromHtml("#FFFFEF00"); } }

        public static XlsxColor CandyAppleRed { get { return FromHtml("#FFFF0800"); } }

        public static XlsxColor CandyPink { get { return FromHtml("#FFE4717A"); } }

        public static XlsxColor CaputMortuum { get { return FromHtml("#FF592720"); } }

        public static XlsxColor Cardinal { get { return FromHtml("#FFC41E3A"); } }

        public static XlsxColor CaribbeanGreen { get { return FromHtml("#FF00CC99"); } }

        public static XlsxColor Carmine { get { return FromHtml("#FF960018"); } }

        public static XlsxColor CarminePink { get { return FromHtml("#FFEB4C42"); } }

        public static XlsxColor CarmineRed { get { return FromHtml("#FFFF0038"); } }

        public static XlsxColor CarnationPink { get { return FromHtml("#FFFFA6C9"); } }

        public static XlsxColor Carnelian { get { return FromHtml("#FFB31B1B"); } }

        public static XlsxColor CarolinaBlue { get { return FromHtml("#FF99BADD"); } }

        public static XlsxColor CarrotOrange { get { return FromHtml("#FFED9121"); } }

        public static XlsxColor Ceil { get { return FromHtml("#FF92A1CF"); } }

        public static XlsxColor Celadon { get { return FromHtml("#FFACE1AF"); } }

        public static XlsxColor CelestialBlue { get { return FromHtml("#FF4997D0"); } }

        public static XlsxColor Cerise { get { return FromHtml("#FFDE3163"); } }

        public static XlsxColor CerisePink { get { return FromHtml("#FFEC3B83"); } }

        public static XlsxColor Cerulean { get { return FromHtml("#FF007BA7"); } }

        public static XlsxColor CeruleanBlue { get { return FromHtml("#FF2A52BE"); } }

        public static XlsxColor Chamoisee { get { return FromHtml("#FFA0785A"); } }

        public static XlsxColor Champagne { get { return FromHtml("#FFF7E7CE"); } }

        public static XlsxColor Charcoal { get { return FromHtml("#FF36454F"); } }

        public static XlsxColor ChartreuseTraditional { get { return FromHtml("#FFDFFF00"); } }

        public static XlsxColor CherryBlossomPink { get { return FromHtml("#FFFFB7C5"); } }

        public static XlsxColor Chocolate1 { get { return FromHtml("#FF7B3F00"); } }

        public static XlsxColor ChromeYellow { get { return FromHtml("#FFFFA700"); } }

        public static XlsxColor Cinereous { get { return FromHtml("#FF98817B"); } }

        public static XlsxColor Cinnabar { get { return FromHtml("#FFE34234"); } }

        public static XlsxColor Citrine { get { return FromHtml("#FFE4D00A"); } }

        public static XlsxColor ClassicRose { get { return FromHtml("#FFFBCCE7"); } }

        public static XlsxColor Cobalt { get { return FromHtml("#FF0047AB"); } }

        public static XlsxColor ColumbiaBlue { get { return FromHtml("#FF9BDDFF"); } }

        public static XlsxColor CoolBlack { get { return FromHtml("#FF002E63"); } }

        public static XlsxColor CoolGrey { get { return FromHtml("#FF8C92AC"); } }

        public static XlsxColor Copper { get { return FromHtml("#FFB87333"); } }

        public static XlsxColor CopperRose { get { return FromHtml("#FF996666"); } }

        public static XlsxColor Coquelicot { get { return FromHtml("#FFFF3800"); } }

        public static XlsxColor CoralPink { get { return FromHtml("#FFF88379"); } }

        public static XlsxColor CoralRed { get { return FromHtml("#FFFF4040"); } }

        public static XlsxColor Cordovan { get { return FromHtml("#FF893F45"); } }

        public static XlsxColor Corn { get { return FromHtml("#FFFBEC5D"); } }

        public static XlsxColor CornellRed { get { return FromHtml("#FFB31B1B"); } }

        public static XlsxColor CosmicLatte { get { return FromHtml("#FFFFF8E7"); } }

        public static XlsxColor CottonCandy { get { return FromHtml("#FFFFBCD9"); } }

        public static XlsxColor Cream { get { return FromHtml("#FFFFFDD0"); } }

        public static XlsxColor CrimsonGlory { get { return FromHtml("#FFBE0032"); } }

        public static XlsxColor CyanProcess { get { return FromHtml("#FF00B7EB"); } }

        public static XlsxColor Daffodil { get { return FromHtml("#FFFFFF31"); } }

        public static XlsxColor Dandelion { get { return FromHtml("#FFF0E130"); } }

        public static XlsxColor DarkBrown { get { return FromHtml("#FF654321"); } }

        public static XlsxColor DarkByzantium { get { return FromHtml("#FF5D3954"); } }

        public static XlsxColor DarkCandyAppleRed { get { return FromHtml("#FFA40000"); } }

        public static XlsxColor DarkCerulean { get { return FromHtml("#FF08457E"); } }

        public static XlsxColor DarkChampagne { get { return FromHtml("#FFC2B280"); } }

        public static XlsxColor DarkChestnut { get { return FromHtml("#FF986960"); } }

        public static XlsxColor DarkCoral { get { return FromHtml("#FFCD5B45"); } }

        public static XlsxColor DarkElectricBlue { get { return FromHtml("#FF536878"); } }

        public static XlsxColor DarkGreen1 { get { return FromHtml("#FF013220"); } }

        public static XlsxColor DarkJungleGreen { get { return FromHtml("#FF1A2421"); } }

        public static XlsxColor DarkLava { get { return FromHtml("#FF483C32"); } }

        public static XlsxColor DarkLavender { get { return FromHtml("#FF734F96"); } }

        public static XlsxColor DarkMidnightBlue { get { return FromHtml("#FF003366"); } }

        public static XlsxColor DarkPastelBlue { get { return FromHtml("#FF779ECB"); } }

        public static XlsxColor DarkPastelGreen { get { return FromHtml("#FF03C03C"); } }

        public static XlsxColor DarkPastelPurple { get { return FromHtml("#FF966FD6"); } }

        public static XlsxColor DarkPastelRed { get { return FromHtml("#FFC23B22"); } }

        public static XlsxColor DarkPink { get { return FromHtml("#FFE75480"); } }

        public static XlsxColor DarkPowderBlue { get { return FromHtml("#FF003399"); } }

        public static XlsxColor DarkRaspberry { get { return FromHtml("#FF872657"); } }

        public static XlsxColor DarkScarlet { get { return FromHtml("#FF560319"); } }

        public static XlsxColor DarkSienna { get { return FromHtml("#FF3C1414"); } }

        public static XlsxColor DarkSpringGreen { get { return FromHtml("#FF177245"); } }

        public static XlsxColor DarkTan { get { return FromHtml("#FF918151"); } }

        public static XlsxColor DarkTangerine { get { return FromHtml("#FFFFA812"); } }

        public static XlsxColor DarkTaupe { get { return FromHtml("#FF483C32"); } }

        public static XlsxColor DarkTerraCotta { get { return FromHtml("#FFCC4E5C"); } }

        public static XlsxColor DartmouthGreen { get { return FromHtml("#FF00693E"); } }

        public static XlsxColor DavysGrey { get { return FromHtml("#FF555555"); } }

        public static XlsxColor DebianRed { get { return FromHtml("#FFD70A53"); } }

        public static XlsxColor DeepCarmine { get { return FromHtml("#FFA9203E"); } }

        public static XlsxColor DeepCarminePink { get { return FromHtml("#FFEF3038"); } }

        public static XlsxColor DeepCarrotOrange { get { return FromHtml("#FFE9692C"); } }

        public static XlsxColor DeepCerise { get { return FromHtml("#FFDA3287"); } }

        public static XlsxColor DeepChampagne { get { return FromHtml("#FFFAD6A5"); } }

        public static XlsxColor DeepChestnut { get { return FromHtml("#FFB94E48"); } }

        public static XlsxColor DeepFuchsia { get { return FromHtml("#FFC154C1"); } }

        public static XlsxColor DeepJungleGreen { get { return FromHtml("#FF004B49"); } }

        public static XlsxColor DeepLilac { get { return FromHtml("#FF9955BB"); } }

        public static XlsxColor DeepMagenta { get { return FromHtml("#FFCC00CC"); } }

        public static XlsxColor DeepPeach { get { return FromHtml("#FFFFCBA4"); } }

        public static XlsxColor DeepSaffron { get { return FromHtml("#FFFF9933"); } }

        public static XlsxColor Denim { get { return FromHtml("#FF1560BD"); } }

        public static XlsxColor Desert { get { return FromHtml("#FFC19A6B"); } }

        public static XlsxColor DesertSand { get { return FromHtml("#FFEDC9AF"); } }

        public static XlsxColor DogwoodRose { get { return FromHtml("#FFD71868"); } }

        public static XlsxColor DollarBill { get { return FromHtml("#FF85BB65"); } }

        public static XlsxColor Drab { get { return FromHtml("#FF967117"); } }

        public static XlsxColor DukeBlue { get { return FromHtml("#FF00009C"); } }

        public static XlsxColor EarthYellow { get { return FromHtml("#FFE1A95F"); } }

        public static XlsxColor Ecru { get { return FromHtml("#FFC2B280"); } }

        public static XlsxColor Eggplant { get { return FromHtml("#FF614051"); } }

        public static XlsxColor Eggshell { get { return FromHtml("#FFF0EAD6"); } }

        public static XlsxColor EgyptianBlue { get { return FromHtml("#FF1034A6"); } }

        public static XlsxColor ElectricBlue { get { return FromHtml("#FF7DF9FF"); } }

        public static XlsxColor ElectricCrimson { get { return FromHtml("#FFFF003F"); } }

        public static XlsxColor ElectricIndigo { get { return FromHtml("#FF6F00FF"); } }

        public static XlsxColor ElectricLavender { get { return FromHtml("#FFF4BBFF"); } }

        public static XlsxColor ElectricLime { get { return FromHtml("#FFCCFF00"); } }

        public static XlsxColor ElectricPurple { get { return FromHtml("#FFBF00FF"); } }

        public static XlsxColor ElectricUltramarine { get { return FromHtml("#FF3F00FF"); } }

        public static XlsxColor ElectricViolet { get { return FromHtml("#FF8F00FF"); } }

        public static XlsxColor Emerald { get { return FromHtml("#FF50C878"); } }

        public static XlsxColor EtonBlue { get { return FromHtml("#FF96C8A2"); } }

        public static XlsxColor Fallow { get { return FromHtml("#FFC19A6B"); } }

        public static XlsxColor FaluRed { get { return FromHtml("#FF801818"); } }

        public static XlsxColor Fandango { get { return FromHtml("#FFB53389"); } }

        public static XlsxColor FashionFuchsia { get { return FromHtml("#FFF400A1"); } }

        public static XlsxColor Fawn { get { return FromHtml("#FFE5AA70"); } }

        public static XlsxColor Feldgrau { get { return FromHtml("#FF4D5D53"); } }

        public static XlsxColor FernGreen { get { return FromHtml("#FF4F7942"); } }

        public static XlsxColor FerrariRed { get { return FromHtml("#FFFF2800"); } }

        public static XlsxColor FieldDrab { get { return FromHtml("#FF6C541E"); } }

        public static XlsxColor FireEngineRed { get { return FromHtml("#FFCE2029"); } }

        public static XlsxColor Flame { get { return FromHtml("#FFE25822"); } }

        public static XlsxColor FlamingoPink { get { return FromHtml("#FFFC8EAC"); } }

        public static XlsxColor Flavescent { get { return FromHtml("#FFF7E98E"); } }

        public static XlsxColor Flax { get { return FromHtml("#FFEEDC82"); } }

        public static XlsxColor FluorescentOrange { get { return FromHtml("#FFFFBF00"); } }

        public static XlsxColor FluorescentYellow { get { return FromHtml("#FFCCFF00"); } }

        public static XlsxColor Folly { get { return FromHtml("#FFFF004F"); } }

        public static XlsxColor ForestGreenTraditional { get { return FromHtml("#FF014421"); } }

        public static XlsxColor FrenchBeige { get { return FromHtml("#FFA67B5B"); } }

        public static XlsxColor FrenchBlue { get { return FromHtml("#FF0072BB"); } }

        public static XlsxColor FrenchLilac { get { return FromHtml("#FF86608E"); } }

        public static XlsxColor FrenchRose { get { return FromHtml("#FFF64A8A"); } }

        public static XlsxColor FuchsiaPink { get { return FromHtml("#FFFF77FF"); } }

        public static XlsxColor Fulvous { get { return FromHtml("#FFE48400"); } }

        public static XlsxColor FuzzyWuzzy { get { return FromHtml("#FFCC6666"); } }

        public static XlsxColor Gamboge { get { return FromHtml("#FFE49B0F"); } }

        public static XlsxColor Ginger { get { return FromHtml("#FFF9F9FF"); } }

        public static XlsxColor Glaucous { get { return FromHtml("#FF6082B6"); } }

        public static XlsxColor GoldenBrown { get { return FromHtml("#FF996515"); } }

        public static XlsxColor GoldenPoppy { get { return FromHtml("#FFFCC200"); } }

        public static XlsxColor GoldenYellow { get { return FromHtml("#FFFFDF00"); } }

        public static XlsxColor GoldMetallic { get { return FromHtml("#FFD4AF37"); } }

        public static XlsxColor GrannySmithApple { get { return FromHtml("#FFA8E4A0"); } }

        public static XlsxColor GrayAsparagus { get { return FromHtml("#FF465945"); } }

        public static XlsxColor GreenPigment { get { return FromHtml("#FF00A550"); } }

        public static XlsxColor GreenRyb { get { return FromHtml("#FF66B032"); } }

        public static XlsxColor Grullo { get { return FromHtml("#FFA99A86"); } }

        public static XlsxColor HalayaUbe { get { return FromHtml("#FF663854"); } }

        public static XlsxColor HanBlue { get { return FromHtml("#FF446CCF"); } }

        public static XlsxColor HanPurple { get { return FromHtml("#FF5218FA"); } }

        public static XlsxColor HansaYellow { get { return FromHtml("#FFE9D66B"); } }

        public static XlsxColor Harlequin { get { return FromHtml("#FF3FFF00"); } }

        public static XlsxColor HarvardCrimson { get { return FromHtml("#FFC90016"); } }

        public static XlsxColor HarvestGold { get { return FromHtml("#FFDA9100"); } }

        public static XlsxColor Heliotrope { get { return FromHtml("#FFDF73FF"); } }

        public static XlsxColor HollywoodCerise { get { return FromHtml("#FFF400A1"); } }

        public static XlsxColor HookersGreen { get { return FromHtml("#FF007000"); } }

        public static XlsxColor HotMagenta { get { return FromHtml("#FFFF1DCE"); } }

        public static XlsxColor HunterGreen { get { return FromHtml("#FF355E3B"); } }

        public static XlsxColor Iceberg { get { return FromHtml("#FF71A6D2"); } }

        public static XlsxColor Icterine { get { return FromHtml("#FFFCF75E"); } }

        public static XlsxColor Inchworm { get { return FromHtml("#FFB2EC5D"); } }

        public static XlsxColor IndiaGreen { get { return FromHtml("#FF138808"); } }

        public static XlsxColor IndianYellow { get { return FromHtml("#FFE3A857"); } }

        public static XlsxColor IndigoDye { get { return FromHtml("#FF00416A"); } }

        public static XlsxColor InternationalKleinBlue { get { return FromHtml("#FF002FA7"); } }

        public static XlsxColor InternationalOrange { get { return FromHtml("#FFFF4F00"); } }

        public static XlsxColor Iris { get { return FromHtml("#FF5A4FCF"); } }

        public static XlsxColor Isabelline { get { return FromHtml("#FFF4F0EC"); } }

        public static XlsxColor IslamicGreen { get { return FromHtml("#FF009000"); } }

        public static XlsxColor Jade { get { return FromHtml("#FF00A86B"); } }

        public static XlsxColor Jasper { get { return FromHtml("#FFD73B3E"); } }

        public static XlsxColor JazzberryJam { get { return FromHtml("#FFA50B5E"); } }

        public static XlsxColor Jonquil { get { return FromHtml("#FFFADA5E"); } }

        public static XlsxColor JuneBud { get { return FromHtml("#FFBDDA57"); } }

        public static XlsxColor JungleGreen { get { return FromHtml("#FF29AB87"); } }

        public static XlsxColor KellyGreen { get { return FromHtml("#FF4CBB17"); } }

        public static XlsxColor KhakiHtmlCssKhaki { get { return FromHtml("#FFC3B091"); } }

        public static XlsxColor LanguidLavender { get { return FromHtml("#FFD6CADD"); } }

        public static XlsxColor LapisLazuli { get { return FromHtml("#FF26619C"); } }

        public static XlsxColor LaSalleGreen { get { return FromHtml("#FF087830"); } }

        public static XlsxColor LaserLemon { get { return FromHtml("#FFFEFE22"); } }

        public static XlsxColor Lava { get { return FromHtml("#FFCF1020"); } }

        public static XlsxColor LavenderBlue { get { return FromHtml("#FFCCCCFF"); } }

        public static XlsxColor LavenderFloral { get { return FromHtml("#FFB57EDC"); } }

        public static XlsxColor LavenderGray { get { return FromHtml("#FFC4C3D0"); } }

        public static XlsxColor LavenderIndigo { get { return FromHtml("#FF9457EB"); } }

        public static XlsxColor LavenderPink { get { return FromHtml("#FFFBAED2"); } }

        public static XlsxColor LavenderPurple { get { return FromHtml("#FF967BB6"); } }

        public static XlsxColor LavenderRose { get { return FromHtml("#FFFBA0E3"); } }

        public static XlsxColor Lemon { get { return FromHtml("#FFFFF700"); } }

        public static XlsxColor LightApricot { get { return FromHtml("#FFFDD5B1"); } }

        public static XlsxColor LightBrown { get { return FromHtml("#FFB5651D"); } }

        public static XlsxColor LightCarminePink { get { return FromHtml("#FFE66771"); } }

        public static XlsxColor LightCornflowerBlue { get { return FromHtml("#FF93CCEA"); } }

        public static XlsxColor LightFuchsiaPink { get { return FromHtml("#FFF984EF"); } }

        public static XlsxColor LightMauve { get { return FromHtml("#FFDCD0FF"); } }

        public static XlsxColor LightPastelPurple { get { return FromHtml("#FFB19CD9"); } }

        public static XlsxColor LightSalmonPink { get { return FromHtml("#FFFF9999"); } }

        public static XlsxColor LightTaupe { get { return FromHtml("#FFB38B6D"); } }

        public static XlsxColor LightThulianPink { get { return FromHtml("#FFE68FAC"); } }

        public static XlsxColor LightYellow1 { get { return FromHtml("#FFFFFFED"); } }

        public static XlsxColor Lilac { get { return FromHtml("#FFC8A2C8"); } }

        public static XlsxColor LimeColorWheel { get { return FromHtml("#FFBFFF00"); } }

        public static XlsxColor LincolnGreen { get { return FromHtml("#FF195905"); } }

        public static XlsxColor Liver { get { return FromHtml("#FF534B4F"); } }

        public static XlsxColor Lust { get { return FromHtml("#FFE62020"); } }

        public static XlsxColor MacaroniAndCheese { get { return FromHtml("#FFFFBD88"); } }

        public static XlsxColor MagentaDye { get { return FromHtml("#FFCA1F7B"); } }

        public static XlsxColor MagentaProcess { get { return FromHtml("#FFFF0090"); } }

        public static XlsxColor MagicMint { get { return FromHtml("#FFAAF0D1"); } }

        public static XlsxColor Magnolia { get { return FromHtml("#FFF8F4FF"); } }

        public static XlsxColor Mahogany { get { return FromHtml("#FFC04000"); } }

        public static XlsxColor Maize { get { return FromHtml("#FFFBEC5D"); } }

        public static XlsxColor MajorelleBlue { get { return FromHtml("#FF6050DC"); } }

        public static XlsxColor Malachite { get { return FromHtml("#FF0BDA51"); } }

        public static XlsxColor Manatee { get { return FromHtml("#FF979AAA"); } }

        public static XlsxColor MangoTango { get { return FromHtml("#FFFF8243"); } }

        public static XlsxColor MaroonX11 { get { return FromHtml("#FFB03060"); } }

        public static XlsxColor Mauve { get { return FromHtml("#FFE0B0FF"); } }

        public static XlsxColor Mauvelous { get { return FromHtml("#FFEF98AA"); } }

        public static XlsxColor MauveTaupe { get { return FromHtml("#FF915F6D"); } }

        public static XlsxColor MayaBlue { get { return FromHtml("#FF73C2FB"); } }

        public static XlsxColor MeatBrown { get { return FromHtml("#FFE5B73B"); } }

        public static XlsxColor MediumAquamarine1 { get { return FromHtml("#FF66DDAA"); } }

        public static XlsxColor MediumCandyAppleRed { get { return FromHtml("#FFE2062C"); } }

        public static XlsxColor MediumCarmine { get { return FromHtml("#FFAF4035"); } }

        public static XlsxColor MediumChampagne { get { return FromHtml("#FFF3E5AB"); } }

        public static XlsxColor MediumElectricBlue { get { return FromHtml("#FF035096"); } }

        public static XlsxColor MediumJungleGreen { get { return FromHtml("#FF1C352D"); } }

        public static XlsxColor MediumPersianBlue { get { return FromHtml("#FF0067A5"); } }

        public static XlsxColor MediumRedViolet { get { return FromHtml("#FFBB3385"); } }

        public static XlsxColor MediumSpringBud { get { return FromHtml("#FFC9DC87"); } }

        public static XlsxColor MediumTaupe { get { return FromHtml("#FF674C47"); } }

        public static XlsxColor Melon { get { return FromHtml("#FFFDBCB4"); } }

        public static XlsxColor MidnightGreenEagleGreen { get { return FromHtml("#FF004953"); } }

        public static XlsxColor MikadoYellow { get { return FromHtml("#FFFFC40C"); } }

        public static XlsxColor Mint { get { return FromHtml("#FF3EB489"); } }

        public static XlsxColor MintGreen { get { return FromHtml("#FF98FF98"); } }

        public static XlsxColor ModeBeige { get { return FromHtml("#FF967117"); } }

        public static XlsxColor MoonstoneBlue { get { return FromHtml("#FF73A9C2"); } }

        public static XlsxColor MordantRed19 { get { return FromHtml("#FFAE0C00"); } }

        public static XlsxColor MossGreen { get { return FromHtml("#FFADDFAD"); } }

        public static XlsxColor MountainMeadow { get { return FromHtml("#FF30BA8F"); } }

        public static XlsxColor MountbattenPink { get { return FromHtml("#FF997A8D"); } }

        public static XlsxColor MsuGreen { get { return FromHtml("#FF18453B"); } }

        public static XlsxColor Mulberry { get { return FromHtml("#FFC54B8C"); } }

        public static XlsxColor Mustard { get { return FromHtml("#FFFFDB58"); } }

        public static XlsxColor Myrtle { get { return FromHtml("#FF21421E"); } }

        public static XlsxColor NadeshikoPink { get { return FromHtml("#FFF6ADC6"); } }

        public static XlsxColor NapierGreen { get { return FromHtml("#FF2A8000"); } }

        public static XlsxColor NaplesYellow { get { return FromHtml("#FFFADA5E"); } }

        public static XlsxColor NeonCarrot { get { return FromHtml("#FFFFA343"); } }

        public static XlsxColor NeonFuchsia { get { return FromHtml("#FFFE59C2"); } }

        public static XlsxColor NeonGreen { get { return FromHtml("#FF39FF14"); } }

        public static XlsxColor NonPhotoBlue { get { return FromHtml("#FFA4DDED"); } }

        public static XlsxColor OceanBoatBlue { get { return FromHtml("#FFCC7422"); } }

        public static XlsxColor Ochre { get { return FromHtml("#FFCC7722"); } }

        public static XlsxColor OldGold { get { return FromHtml("#FFCFB53B"); } }

        public static XlsxColor OldLavender { get { return FromHtml("#FF796878"); } }

        public static XlsxColor OldMauve { get { return FromHtml("#FF673147"); } }

        public static XlsxColor OldRose { get { return FromHtml("#FFC08081"); } }

        public static XlsxColor OliveDrab7 { get { return FromHtml("#FF3C341F"); } }

        public static XlsxColor Olivine { get { return FromHtml("#FF9AB973"); } }

        public static XlsxColor Onyx { get { return FromHtml("#FF0F0F0F"); } }

        public static XlsxColor OperaMauve { get { return FromHtml("#FFB784A7"); } }

        public static XlsxColor OrangeColorWheel { get { return FromHtml("#FFFF7F00"); } }

        public static XlsxColor OrangePeel { get { return FromHtml("#FFFF9F00"); } }

        public static XlsxColor OrangeRyb { get { return FromHtml("#FFFB9902"); } }

        public static XlsxColor OtterBrown { get { return FromHtml("#FF654321"); } }

        public static XlsxColor OuCrimsonRed { get { return FromHtml("#FF990000"); } }

        public static XlsxColor OuterSpace { get { return FromHtml("#FF414A4C"); } }

        public static XlsxColor OutrageousOrange { get { return FromHtml("#FFFF6E4A"); } }

        public static XlsxColor OxfordBlue { get { return FromHtml("#FF002147"); } }

        public static XlsxColor PakistanGreen { get { return FromHtml("#FF00421B"); } }

        public static XlsxColor PalatinateBlue { get { return FromHtml("#FF273BE2"); } }

        public static XlsxColor PalatinatePurple { get { return FromHtml("#FF682860"); } }

        public static XlsxColor PaleAqua { get { return FromHtml("#FFBCD4E6"); } }

        public static XlsxColor PaleBrown { get { return FromHtml("#FF987654"); } }

        public static XlsxColor PaleCarmine { get { return FromHtml("#FFAF4035"); } }

        public static XlsxColor PaleCerulean { get { return FromHtml("#FF9BC4E2"); } }

        public static XlsxColor PaleChestnut { get { return FromHtml("#FFDDADAF"); } }

        public static XlsxColor PaleCopper { get { return FromHtml("#FFDA8A67"); } }

        public static XlsxColor PaleCornflowerBlue { get { return FromHtml("#FFABCDEF"); } }

        public static XlsxColor PaleGold { get { return FromHtml("#FFE6BE8A"); } }

        public static XlsxColor PaleMagenta { get { return FromHtml("#FFF984E5"); } }

        public static XlsxColor PalePink { get { return FromHtml("#FFFADADD"); } }

        public static XlsxColor PaleRobinEggBlue { get { return FromHtml("#FF96DED1"); } }

        public static XlsxColor PaleSilver { get { return FromHtml("#FFC9C0BB"); } }

        public static XlsxColor PaleSpringBud { get { return FromHtml("#FFECEBBD"); } }

        public static XlsxColor PaleTaupe { get { return FromHtml("#FFBC987E"); } }

        public static XlsxColor PansyPurple { get { return FromHtml("#FF78184A"); } }

        public static XlsxColor ParisGreen { get { return FromHtml("#FF50C878"); } }

        public static XlsxColor PastelBlue { get { return FromHtml("#FFAEC6CF"); } }

        public static XlsxColor PastelBrown { get { return FromHtml("#FF836953"); } }

        public static XlsxColor PastelGray { get { return FromHtml("#FFCFCFC4"); } }

        public static XlsxColor PastelGreen { get { return FromHtml("#FF77DD77"); } }

        public static XlsxColor PastelMagenta { get { return FromHtml("#FFF49AC2"); } }

        public static XlsxColor PastelOrange { get { return FromHtml("#FFFFB347"); } }

        public static XlsxColor PastelPink { get { return FromHtml("#FFFFD1DC"); } }

        public static XlsxColor PastelPurple { get { return FromHtml("#FFB39EB5"); } }

        public static XlsxColor PastelRed { get { return FromHtml("#FFFF6961"); } }

        public static XlsxColor PastelViolet { get { return FromHtml("#FFCB99C9"); } }

        public static XlsxColor PastelYellow { get { return FromHtml("#FFFDFD96"); } }

        public static XlsxColor PaynesGrey { get { return FromHtml("#FF40404F"); } }

        public static XlsxColor Peach { get { return FromHtml("#FFFFE5B4"); } }

        public static XlsxColor PeachOrange { get { return FromHtml("#FFFFCC99"); } }

        public static XlsxColor PeachYellow { get { return FromHtml("#FFFADFAD"); } }

        public static XlsxColor Pear { get { return FromHtml("#FFD1E231"); } }

        public static XlsxColor Pearl { get { return FromHtml("#FFF0EAD6"); } }

        public static XlsxColor Peridot { get { return FromHtml("#FFE6E200"); } }

        public static XlsxColor Periwinkle { get { return FromHtml("#FFCCCCFF"); } }

        public static XlsxColor PersianBlue { get { return FromHtml("#FF1C39BB"); } }

        public static XlsxColor PersianGreen { get { return FromHtml("#FF00A693"); } }

        public static XlsxColor PersianIndigo { get { return FromHtml("#FF32127A"); } }

        public static XlsxColor PersianOrange { get { return FromHtml("#FFD99058"); } }

        public static XlsxColor PersianPink { get { return FromHtml("#FFF77FBE"); } }

        public static XlsxColor PersianPlum { get { return FromHtml("#FF701C1C"); } }

        public static XlsxColor PersianRed { get { return FromHtml("#FFCC3333"); } }

        public static XlsxColor PersianRose { get { return FromHtml("#FFFE28A2"); } }

        public static XlsxColor Persimmon { get { return FromHtml("#FFEC5800"); } }

        public static XlsxColor Phlox { get { return FromHtml("#FFDF00FF"); } }

        public static XlsxColor PhthaloBlue { get { return FromHtml("#FF000F89"); } }

        public static XlsxColor PhthaloGreen { get { return FromHtml("#FF123524"); } }

        public static XlsxColor PiggyPink { get { return FromHtml("#FFFDDDE6"); } }

        public static XlsxColor PineGreen { get { return FromHtml("#FF01796F"); } }

        public static XlsxColor PinkOrange { get { return FromHtml("#FFFF9966"); } }

        public static XlsxColor PinkPearl { get { return FromHtml("#FFE7ACCF"); } }

        public static XlsxColor PinkSherbet { get { return FromHtml("#FFF78FA7"); } }

        public static XlsxColor Pistachio { get { return FromHtml("#FF93C572"); } }

        public static XlsxColor Platinum { get { return FromHtml("#FFE5E4E2"); } }

        public static XlsxColor PlumTraditional { get { return FromHtml("#FF8E4585"); } }

        public static XlsxColor PortlandOrange { get { return FromHtml("#FFFF5A36"); } }

        public static XlsxColor PrincetonOrange { get { return FromHtml("#FFFF8F00"); } }

        public static XlsxColor Prune { get { return FromHtml("#FF701C1C"); } }

        public static XlsxColor PrussianBlue { get { return FromHtml("#FF003153"); } }

        public static XlsxColor PsychedelicPurple { get { return FromHtml("#FFDF00FF"); } }

        public static XlsxColor Puce { get { return FromHtml("#FFCC8899"); } }

        public static XlsxColor Pumpkin { get { return FromHtml("#FFFF7518"); } }

        public static XlsxColor PurpleHeart { get { return FromHtml("#FF69359C"); } }

        public static XlsxColor PurpleMountainMajesty { get { return FromHtml("#FF9678B6"); } }

        public static XlsxColor PurpleMunsell { get { return FromHtml("#FF9F00C5"); } }

        public static XlsxColor PurplePizzazz { get { return FromHtml("#FFFE4EDA"); } }

        public static XlsxColor PurpleTaupe { get { return FromHtml("#FF50404D"); } }

        public static XlsxColor PurpleX11 { get { return FromHtml("#FFA020F0"); } }

        public static XlsxColor RadicalRed { get { return FromHtml("#FFFF355E"); } }

        public static XlsxColor Raspberry { get { return FromHtml("#FFE30B5D"); } }

        public static XlsxColor RaspberryGlace { get { return FromHtml("#FF915F6D"); } }

        public static XlsxColor RaspberryPink { get { return FromHtml("#FFE25098"); } }

        public static XlsxColor RaspberryRose { get { return FromHtml("#FFB3446C"); } }

        public static XlsxColor RawUmber { get { return FromHtml("#FF826644"); } }

        public static XlsxColor RazzleDazzleRose { get { return FromHtml("#FFFF33CC"); } }

        public static XlsxColor Razzmatazz { get { return FromHtml("#FFE3256B"); } }

        public static XlsxColor RedMunsell { get { return FromHtml("#FFF2003C"); } }

        public static XlsxColor RedNcs { get { return FromHtml("#FFC40233"); } }

        public static XlsxColor RedPigment { get { return FromHtml("#FFED1C24"); } }

        public static XlsxColor RedRyb { get { return FromHtml("#FFFE2712"); } }

        public static XlsxColor Redwood { get { return FromHtml("#FFAB4E52"); } }

        public static XlsxColor Regalia { get { return FromHtml("#FF522D80"); } }

        public static XlsxColor RichBlack { get { return FromHtml("#FF004040"); } }

        public static XlsxColor RichBrilliantLavender { get { return FromHtml("#FFF1A7FE"); } }

        public static XlsxColor RichCarmine { get { return FromHtml("#FFD70040"); } }

        public static XlsxColor RichElectricBlue { get { return FromHtml("#FF0892D0"); } }

        public static XlsxColor RichLavender { get { return FromHtml("#FFA76BCF"); } }

        public static XlsxColor RichLilac { get { return FromHtml("#FFB666D2"); } }

        public static XlsxColor RichMaroon { get { return FromHtml("#FFB03060"); } }

        public static XlsxColor RifleGreen { get { return FromHtml("#FF414833"); } }

        public static XlsxColor RobinEggBlue { get { return FromHtml("#FF00CCCC"); } }

        public static XlsxColor Rose { get { return FromHtml("#FFFF007F"); } }

        public static XlsxColor RoseBonbon { get { return FromHtml("#FFF9429E"); } }

        public static XlsxColor RoseEbony { get { return FromHtml("#FF674846"); } }

        public static XlsxColor RoseGold { get { return FromHtml("#FFB76E79"); } }

        public static XlsxColor RoseMadder { get { return FromHtml("#FFE32636"); } }

        public static XlsxColor RosePink { get { return FromHtml("#FFFF66CC"); } }

        public static XlsxColor RoseQuartz { get { return FromHtml("#FFAA98A9"); } }

        public static XlsxColor RoseTaupe { get { return FromHtml("#FF905D5D"); } }

        public static XlsxColor RoseVale { get { return FromHtml("#FFAB4E52"); } }

        public static XlsxColor Rosewood { get { return FromHtml("#FF65000B"); } }

        public static XlsxColor RossoCorsa { get { return FromHtml("#FFD40000"); } }

        public static XlsxColor RoyalAzure { get { return FromHtml("#FF0038A8"); } }

        public static XlsxColor RoyalBlueTraditional { get { return FromHtml("#FF002366"); } }

        public static XlsxColor RoyalFuchsia { get { return FromHtml("#FFCA2C92"); } }

        public static XlsxColor RoyalPurple { get { return FromHtml("#FF7851A9"); } }

        public static XlsxColor Ruby { get { return FromHtml("#FFE0115F"); } }

        public static XlsxColor Ruddy { get { return FromHtml("#FFFF0028"); } }

        public static XlsxColor RuddyBrown { get { return FromHtml("#FFBB6528"); } }

        public static XlsxColor RuddyPink { get { return FromHtml("#FFE18E96"); } }

        public static XlsxColor Rufous { get { return FromHtml("#FFA81C07"); } }

        public static XlsxColor Russet { get { return FromHtml("#FF80461B"); } }

        public static XlsxColor Rust { get { return FromHtml("#FFB7410E"); } }

        public static XlsxColor SacramentoStateGreen { get { return FromHtml("#FF00563F"); } }

        public static XlsxColor SafetyOrangeBlazeOrange { get { return FromHtml("#FFFF6700"); } }

        public static XlsxColor Saffron { get { return FromHtml("#FFF4C430"); } }

        public static XlsxColor Salmon1 { get { return FromHtml("#FFFF8C69"); } }

        public static XlsxColor SalmonPink { get { return FromHtml("#FFFF91A4"); } }

        public static XlsxColor Sand { get { return FromHtml("#FFC2B280"); } }

        public static XlsxColor SandDune { get { return FromHtml("#FF967117"); } }

        public static XlsxColor Sandstorm { get { return FromHtml("#FFECD540"); } }

        public static XlsxColor SandyTaupe { get { return FromHtml("#FF967117"); } }

        public static XlsxColor Sangria { get { return FromHtml("#FF92000A"); } }

        public static XlsxColor SapGreen { get { return FromHtml("#FF507D2A"); } }

        public static XlsxColor Sapphire { get { return FromHtml("#FF082567"); } }

        public static XlsxColor SatinSheenGold { get { return FromHtml("#FFCBA135"); } }

        public static XlsxColor Scarlet { get { return FromHtml("#FFFF2000"); } }

        public static XlsxColor SchoolBusYellow { get { return FromHtml("#FFFFD800"); } }

        public static XlsxColor ScreaminGreen { get { return FromHtml("#FF76FF7A"); } }

        public static XlsxColor SealBrown { get { return FromHtml("#FF321414"); } }

        public static XlsxColor SelectiveYellow { get { return FromHtml("#FFFFBA00"); } }

        public static XlsxColor Sepia { get { return FromHtml("#FF704214"); } }

        public static XlsxColor Shadow { get { return FromHtml("#FF8A795D"); } }

        public static XlsxColor ShamrockGreen { get { return FromHtml("#FF009E60"); } }

        public static XlsxColor ShockingPink { get { return FromHtml("#FFFC0FC0"); } }

        public static XlsxColor Sienna1 { get { return FromHtml("#FF882D17"); } }

        public static XlsxColor Sinopia { get { return FromHtml("#FFCB410B"); } }

        public static XlsxColor Skobeloff { get { return FromHtml("#FF007474"); } }

        public static XlsxColor SkyMagenta { get { return FromHtml("#FFCF71AF"); } }

        public static XlsxColor SmaltDarkPowderBlue { get { return FromHtml("#FF003399"); } }

        public static XlsxColor SmokeyTopaz { get { return FromHtml("#FF933D41"); } }

        public static XlsxColor SmokyBlack { get { return FromHtml("#FF100C08"); } }

        public static XlsxColor SpiroDiscoBall { get { return FromHtml("#FF0FC0FC"); } }

        public static XlsxColor SplashedWhite { get { return FromHtml("#FFFEFDFF"); } }

        public static XlsxColor SpringBud { get { return FromHtml("#FFA7FC00"); } }

        public static XlsxColor StPatricksBlue { get { return FromHtml("#FF23297A"); } }

        public static XlsxColor StilDeGrainYellow { get { return FromHtml("#FFFADA5E"); } }

        public static XlsxColor Straw { get { return FromHtml("#FFE4D96F"); } }

        public static XlsxColor Sunglow { get { return FromHtml("#FFFFCC33"); } }

        public static XlsxColor Sunset { get { return FromHtml("#FFFAD6A5"); } }

        public static XlsxColor Tangelo { get { return FromHtml("#FFF94D00"); } }

        public static XlsxColor Tangerine { get { return FromHtml("#FFF28500"); } }

        public static XlsxColor TangerineYellow { get { return FromHtml("#FFFFCC00"); } }

        public static XlsxColor Taupe { get { return FromHtml("#FF483C32"); } }

        public static XlsxColor TaupeGray { get { return FromHtml("#FF8B8589"); } }

        public static XlsxColor TeaGreen { get { return FromHtml("#FFD0F0C0"); } }

        public static XlsxColor TealBlue { get { return FromHtml("#FF367588"); } }

        public static XlsxColor TealGreen { get { return FromHtml("#FF006D5B"); } }

        public static XlsxColor TeaRoseOrange { get { return FromHtml("#FFF88379"); } }

        public static XlsxColor TeaRoseRose { get { return FromHtml("#FFF4C2C2"); } }

        public static XlsxColor TennéTawny { get { return FromHtml("#FFCD5700"); } }

        public static XlsxColor TerraCotta { get { return FromHtml("#FFE2725B"); } }

        public static XlsxColor ThulianPink { get { return FromHtml("#FFDE6FA1"); } }

        public static XlsxColor TickleMePink { get { return FromHtml("#FFFC89AC"); } }

        public static XlsxColor TiffanyBlue { get { return FromHtml("#FF0ABAB5"); } }

        public static XlsxColor TigersEye { get { return FromHtml("#FFE08D3C"); } }

        public static XlsxColor Timberwolf { get { return FromHtml("#FFDBD7D2"); } }

        public static XlsxColor TitaniumYellow { get { return FromHtml("#FFEEE600"); } }

        public static XlsxColor Toolbox { get { return FromHtml("#FF746CC0"); } }

        public static XlsxColor TractorRed { get { return FromHtml("#FFFD0E35"); } }

        public static XlsxColor TropicalRainForest { get { return FromHtml("#FF00755E"); } }

        public static XlsxColor TuftsBlue { get { return FromHtml("#FF417DC1"); } }

        public static XlsxColor Tumbleweed { get { return FromHtml("#FFDEAA88"); } }

        public static XlsxColor TurkishRose { get { return FromHtml("#FFB57281"); } }

        public static XlsxColor Turquoise1 { get { return FromHtml("#FF30D5C8"); } }

        public static XlsxColor TurquoiseBlue { get { return FromHtml("#FF00FFEF"); } }

        public static XlsxColor TurquoiseGreen { get { return FromHtml("#FFA0D6B4"); } }

        public static XlsxColor TuscanRed { get { return FromHtml("#FF823535"); } }

        public static XlsxColor TwilightLavender { get { return FromHtml("#FF8A496B"); } }

        public static XlsxColor TyrianPurple { get { return FromHtml("#FF66023C"); } }

        public static XlsxColor UaBlue { get { return FromHtml("#FF0033AA"); } }

        public static XlsxColor UaRed { get { return FromHtml("#FFD9004C"); } }

        public static XlsxColor Ube { get { return FromHtml("#FF8878C3"); } }

        public static XlsxColor UclaBlue { get { return FromHtml("#FF536895"); } }

        public static XlsxColor UclaGold { get { return FromHtml("#FFFFB300"); } }

        public static XlsxColor UfoGreen { get { return FromHtml("#FF3CD070"); } }

        public static XlsxColor Ultramarine { get { return FromHtml("#FF120A8F"); } }

        public static XlsxColor UltramarineBlue { get { return FromHtml("#FF4166F5"); } }

        public static XlsxColor UltraPink { get { return FromHtml("#FFFF6FFF"); } }

        public static XlsxColor Umber { get { return FromHtml("#FF635147"); } }

        public static XlsxColor UnitedNationsBlue { get { return FromHtml("#FF5B92E5"); } }

        public static XlsxColor UnmellowYellow { get { return FromHtml("#FFFFFF66"); } }

        public static XlsxColor UpForestGreen { get { return FromHtml("#FF014421"); } }

        public static XlsxColor UpMaroon { get { return FromHtml("#FF7B1113"); } }

        public static XlsxColor UpsdellRed { get { return FromHtml("#FFAE2029"); } }

        public static XlsxColor Urobilin { get { return FromHtml("#FFE1AD21"); } }

        public static XlsxColor UscCardinal { get { return FromHtml("#FF990000"); } }

        public static XlsxColor UscGold { get { return FromHtml("#FFFFCC00"); } }

        public static XlsxColor UtahCrimson { get { return FromHtml("#FFD3003F"); } }

        public static XlsxColor Vanilla { get { return FromHtml("#FFF3E5AB"); } }

        public static XlsxColor VegasGold { get { return FromHtml("#FFC5B358"); } }

        public static XlsxColor VenetianRed { get { return FromHtml("#FFC80815"); } }

        public static XlsxColor Verdigris { get { return FromHtml("#FF43B3AE"); } }

        public static XlsxColor Vermilion { get { return FromHtml("#FFE34234"); } }

        public static XlsxColor Veronica { get { return FromHtml("#FFA020F0"); } }

        public static XlsxColor Violet1 { get { return FromHtml("#FF8F00FF"); } }

        public static XlsxColor VioletColorWheel { get { return FromHtml("#FF7F00FF"); } }

        public static XlsxColor VioletRyb { get { return FromHtml("#FF8601AF"); } }

        public static XlsxColor Viridian { get { return FromHtml("#FF40826D"); } }

        public static XlsxColor VividAuburn { get { return FromHtml("#FF922724"); } }

        public static XlsxColor VividBurgundy { get { return FromHtml("#FF9F1D35"); } }

        public static XlsxColor VividCerise { get { return FromHtml("#FFDA1D81"); } }

        public static XlsxColor VividTangerine { get { return FromHtml("#FFFFA089"); } }

        public static XlsxColor VividViolet { get { return FromHtml("#FF9F00FF"); } }

        public static XlsxColor WarmBlack { get { return FromHtml("#FF004242"); } }

        public static XlsxColor Wenge { get { return FromHtml("#FF645452"); } }

        public static XlsxColor WildBlueYonder { get { return FromHtml("#FFA2ADD0"); } }

        public static XlsxColor WildStrawberry { get { return FromHtml("#FFFF43A4"); } }

        public static XlsxColor WildWatermelon { get { return FromHtml("#FFFC6C85"); } }

        public static XlsxColor Wisteria { get { return FromHtml("#FFC9A0DC"); } }

        public static XlsxColor Xanadu { get { return FromHtml("#FF738678"); } }

        public static XlsxColor YaleBlue { get { return FromHtml("#FF0F4D92"); } }

        public static XlsxColor YellowMunsell { get { return FromHtml("#FFEFCC00"); } }

        public static XlsxColor YellowNcs { get { return FromHtml("#FFFFD300"); } }

        public static XlsxColor YellowProcess { get { return FromHtml("#FFFFEF00"); } }

        public static XlsxColor YellowRyb { get { return FromHtml("#FFFEFE33"); } }

        public static XlsxColor Zaffre { get { return FromHtml("#FF0014A8"); } }

        public static XlsxColor ZinnwalditeBrown { get { return FromHtml("#FF2C1608"); } }

        public static XlsxColor Transparent { get { return FromColor(Color.Transparent); } }
    }
}
