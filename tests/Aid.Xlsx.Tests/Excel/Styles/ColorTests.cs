using System.Globalization;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Aid.Xlsx.Tests.Excel
{
    [TestFixture]
    public class ColorTests
    {
        [Test]
        public void ColorEqualOperatorInPlace()
        {
            Assert.IsTrue(XlsxColor.Black == XlsxColor.Black);
        }

        [Test]
        public void ColorNotEqualOperatorInPlace()
        {
            Assert.IsFalse(XlsxColor.Black != XlsxColor.Black);
        }

        [Test]
        public void ColorNamedVsHTML()
        {
            Assert.IsTrue(XlsxColor.Black == XlsxColor.FromHtml("#000000"));
        }

        [Test]
        public void DefaultStyleColorIsAutomatic()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(XlsxColor.Automatic, ws.FirstCell().Style.Fill.BackgroundColor);
        }

        [Test]
        public void AutomaticColorCantBeResolvedToColor()
        {
            Assert.That(() => _ = XlsxColor.Automatic.Color, Throws.TypeOf<InvalidOperationException>().With.Message.EqualTo("Cannot convert automatic color to Color."));
        }

        [Test]
        public void CanConvertXLColorToColorType()
        {
            var xlColor1 = XlsxColor.Red;
            var xlColor2 = XlsxColor.FromIndex(20);
            var xlColor3 = XlsxColor.FromTheme(XlsxThemeColor.Accent1);
            var xlColor4 = XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.4);

            var color1 = new ForegroundColor().FromXlsxColor<ForegroundColor>(xlColor1);
            var color2 = new ForegroundColor().FromXlsxColor<ForegroundColor>(xlColor2);
            var color3 = new BackgroundColor().FromXlsxColor<BackgroundColor>(xlColor3);
            var color4 = new BackgroundColor().FromXlsxColor<BackgroundColor>(xlColor4);

            Assert.AreEqual("FFFF0000", color1.Rgb.Value);
            Assert.IsNull(color1.Indexed);
            Assert.IsNull(color1.Theme);
            Assert.IsNull(color1.Tint);

            Assert.IsNull(color2.Rgb);
            Assert.AreEqual(20, color2.Indexed.Value);
            Assert.IsNull(color2.Theme);
            Assert.IsNull(color2.Tint);

            Assert.IsNull(color3.Rgb);
            Assert.IsNull(color3.Indexed);
            Assert.AreEqual(4, color3.Theme.Value);
            Assert.IsNull(color3.Tint);

            Assert.IsNull(color4.Rgb);
            Assert.IsNull(color4.Indexed);
            Assert.AreEqual(5, color4.Theme.Value);
            Assert.AreEqual(0.4, color4.Tint.Value);
        }

        [Test]
        public void CanConvertXlColorToX14ColorType()
        {
            var xlColor1 = XlsxColor.Red;
            var xlColor2 = XlsxColor.FromIndex(20);
            var xlColor3 = XlsxColor.FromTheme(XlsxThemeColor.Accent1);
            var xlColor4 = XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.4);

            var color1 = new X14.AxisColor().FromXlsxColor<X14.AxisColor>(xlColor1);
            var color2 = new X14.BorderColor().FromXlsxColor<X14.BorderColor>(xlColor2);
            var color3 = new X14.FillColor().FromXlsxColor<X14.FillColor>(xlColor3);
            var color4 = new X14.HighMarkerColor().FromXlsxColor<X14.HighMarkerColor>(xlColor4);

            Assert.AreEqual("FFFF0000", color1.Rgb.Value);
            Assert.IsNull(color1.Indexed);
            Assert.IsNull(color1.Theme);
            Assert.IsNull(color1.Tint);

            Assert.IsNull(color2.Rgb);
            Assert.AreEqual(20, color2.Indexed.Value);
            Assert.IsNull(color2.Theme);
            Assert.IsNull(color2.Tint);

            Assert.IsNull(color3.Rgb);
            Assert.IsNull(color3.Indexed);
            Assert.AreEqual(4, color3.Theme.Value);
            Assert.IsNull(color3.Tint);

            Assert.IsNull(color4.Rgb);
            Assert.IsNull(color4.Indexed);
            Assert.AreEqual(5, color4.Theme.Value);
            Assert.AreEqual(0.4, color4.Tint.Value);
        }

        [Test]
        public void CanConvertColorTypeToXlColor()
        {
            var color1 = new ForegroundColor { Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000") };
            var color2 = new ForegroundColor { Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20) };
            var color3 = new BackgroundColor { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
            var color4 = new BackgroundColor
            {
                Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
                Tint = new DocumentFormat.OpenXml.DoubleValue(0.4)
            };

            var xlColor1 = color1.ToXlsxColor();
            var xlColor2 = color2.ToXlsxColor();
            var xlColor3 = color3.ToXlsxColor();
            var xlColor4 = color4.ToXlsxColor();

            Assert.AreEqual(XlsxColorType.Color, xlColor1.ColorType);
            Assert.AreEqual(XlsxColor.Red.Color, xlColor1.Color);

            Assert.AreEqual(XlsxColorType.Indexed, xlColor2.ColorType);
            Assert.AreEqual(20, xlColor2.Indexed);

            Assert.AreEqual(XlsxColorType.Theme, xlColor3.ColorType);
            Assert.AreEqual(XlsxThemeColor.Accent1, xlColor3.ThemeColor);
            Assert.AreEqual(0, xlColor3.ThemeTint, XlsxHelper.Epsilon);

            Assert.AreEqual(XlsxColorType.Theme, xlColor4.ColorType);
            Assert.AreEqual(XlsxThemeColor.Accent1, xlColor4.ThemeColor);
            Assert.AreEqual(0.4, xlColor4.ThemeTint, XlsxHelper.Epsilon);
        }

        [Test]
        public void CanConvertX14ColorTypeToXlColor()
        {
            var color1 = new X14.AxisColor { Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000") };
            var color2 = new X14.BorderColor { Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20) };
            var color3 = new X14.FillColor { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
            var color4 = new X14.HighMarkerColor
            {
                Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
                Tint = new DocumentFormat.OpenXml.DoubleValue(0.4)
            };

            var xlColor1 = color1.ToXlsxColor();
            var xlColor2 = color2.ToXlsxColor();
            var xlColor3 = color3.ToXlsxColor();
            var xlColor4 = color4.ToXlsxColor();

            Assert.AreEqual(XlsxColorType.Color, xlColor1.ColorType);
            Assert.AreEqual(XlsxColor.Red.Color, xlColor1.Color);

            Assert.AreEqual(XlsxColorType.Indexed, xlColor2.ColorType);
            Assert.AreEqual(20, xlColor2.Indexed);

            Assert.AreEqual(XlsxColorType.Theme, xlColor3.ColorType);
            Assert.AreEqual(XlsxThemeColor.Accent1, xlColor3.ThemeColor);
            Assert.AreEqual(0, xlColor3.ThemeTint, XlsxHelper.Epsilon);

            Assert.AreEqual(XlsxColorType.Theme, xlColor4.ColorType);
            Assert.AreEqual(XlsxThemeColor.Accent1, xlColor4.ThemeColor);
            Assert.AreEqual(0.4, xlColor4.ThemeTint, XlsxHelper.Epsilon);
        }

        [Test]
        public void CanParseColorWithHashAsCultureLineSeparator()
        {
            // upstream issue #675
            var culture = CultureInfo.CreateSpecificCulture("en-US");
            culture.TextInfo.ListSeparator = "#";
            Thread.CurrentThread.CurrentCulture = culture;
            var color = XlsxColor.FromHtml("#FF008000");
            Assert.AreEqual(XlsxColor.Green, color);
        }

        [TestCaseSource(nameof(ToStringTestCases))]
        public void ToString_works_for_all_color_types(XlsxColor colorType, string expectedString)
        {
            Assert.AreEqual(expectedString, colorType.ToString());
        }

        private static IEnumerable<TestCaseData<XlsxColor, string>> ToStringTestCases()
        {
            yield return new TestCaseData<XlsxColor, string>(XlsxColor.FromArgb(0xFF804010), "FF804010");
            yield return new TestCaseData<XlsxColor, string>(XlsxColor.FromTheme(XlsxThemeColor.Text1, 0.25), "Color Theme: Text1, Tint: 0.25");
            yield return new TestCaseData<XlsxColor, string>(XlsxColor.FromIndex(14), "Color Index: 14");
            yield return new TestCaseData<XlsxColor, string>(XlsxColor.Automatic, "Automatic");
        }
    }
}
