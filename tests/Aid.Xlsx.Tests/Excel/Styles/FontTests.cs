using Aid.Xlsx.Cells;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles.Formatting;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Tests.Utils;
using Aid.Xlsx.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.Styles
{
    public class FontTests
    {
        private readonly XlsxFontFormatValue _defaultFormat = XlsxFontFormatValue.Default;

        [Test]
        public void XlsxFontFormatValue_GetHashCode_IsCaseInsensitive()
        {
            var fontKey1 = _defaultFormat with { Name = "Arial" };
            var fontKey2 = _defaultFormat with { Name = "Times New Roman" };
            var fontKey3 = _defaultFormat with { Name = "TIMES NEW ROMAN" };

            Assert.AreNotEqual(fontKey1.GetHashCode(), fontKey2.GetHashCode());
            Assert.AreEqual(fontKey2.GetHashCode(), fontKey3.GetHashCode());
        }

        [Test]
        public void XlsxFontFormatValue_Equals_IsCaseInsensitive()
        {
            var fontKey1 = _defaultFormat with { Name = "Arial" };
            var fontKey2 = _defaultFormat with { Name = "Times New Roman" };
            var fontKey3 = _defaultFormat with { Name = "TIMES NEW ROMAN" };

            Assert.IsFalse(fontKey1.Equals(fontKey2));
            Assert.IsTrue(fontKey2.Equals(fontKey3));
        }

        [Test]
        [TestCaseSource(nameof(FontApiSetters))]
        public void Font_property_can_be_individually_set(FormatTestCase<IXlsxFont> testCase)
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            var cellsFormat = ((XlsxCells)ws.Cells("A1:C4")).Format;
            var cellFormat = ((XlsxCell)ws.Cell("B2")).Format;

            foreach (var testValue in testCase.Values)
            {
                testCase.SetPropertyValue(cellsFormat.Font, testValue);
                var setValue = testCase.GetPropertyValue(cellFormat.Font);
                Assert.AreEqual(testValue, setValue);
            }
        }

        [Test]
        [TestCaseSource(nameof(FontApiSetters))]
        public void Dxf_font_property_can_be_individually_set(FormatTestCase<IXlsxFont> testCase)
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            var cf = ws.AddConditionalFormat();

            foreach (var testValue in testCase.Values)
            {
                testCase.SetPropertyValue(cf.Style.Font, testValue);
                var setValue = testCase.GetPropertyValue(cf.Style.Font);
                Assert.AreEqual(testValue, setValue);
            }
        }

        [Test]
        public void Font_can_be_set_by_assigning_font()
        {
            // Arrange
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Style
                .Font.SetBold()
                .Font.SetItalic()
                .Font.SetUnderline(XlsxFontUnderlineValues.DoubleAccounting)
                .Font.SetStrikethrough()
                .Font.SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues.Superscript)
                .Font.SetShadow()
                .Font.SetFontSize(25)
                .Font.SetFontColor(XlsxColor.Red)
                .Font.SetFontName("Arial")
                .Font.SetFontFamilyNumbering(XlsxFontFamilyNumberingValues.Decorative)
                .Font.SetFontCharSet(XlsxFontCharSet.Hangul)
                .Font.SetFontScheme(XlsxFontScheme.Minor);

            // Act
            ws.Cell("A2").Style.Font = ws.Cell("A1").Style.Font;

            // Assert
            var copiedFont = ws.Cell("A2").Style.Font;
            Assert.IsTrue(copiedFont.Bold);
            Assert.IsTrue(copiedFont.Italic);
            Assert.AreEqual(XlsxFontUnderlineValues.DoubleAccounting, copiedFont.Underline);
            Assert.IsTrue(copiedFont.Strikethrough);
            Assert.AreEqual(XlsxFontVerticalTextAlignmentValues.Superscript, copiedFont.VerticalAlignment);
            Assert.IsTrue(copiedFont.Shadow);
            Assert.AreEqual(25, copiedFont.FontSize);
            Assert.AreEqual(XlsxColor.Red, copiedFont.FontColor);
            Assert.AreEqual("Arial", copiedFont.FontName);
            Assert.AreEqual(XlsxFontFamilyNumberingValues.Decorative, copiedFont.FontFamilyNumbering);
            Assert.AreEqual(XlsxFontCharSet.Hangul, copiedFont.FontCharSet);
            Assert.AreEqual(XlsxFontScheme.Minor, copiedFont.FontScheme);
        }

        [Test]
        public void Font_can_be_checked_for_equality()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            var testFont = ws.Cell("A1").Style.Font;
            var equalFont = ws.Cell("A2").Style.Font;

            Assert.AreEqual(testFont, equalFont);
            var makeDifferentFont = new Action<IXlsxFont>[]
            {
                x => x.Bold = !x.Bold,
                x => x.Italic = !x.Italic,
                x => x.Underline = XlsxFontUnderlineValues.DoubleAccounting,
                x => x.Strikethrough = !x.Strikethrough,
                x => x.VerticalAlignment = XlsxFontVerticalTextAlignmentValues.Superscript,
                x => x.Shadow = !x.Shadow,
                x => x.FontSize = 25,
                x => x.FontColor = XlsxColor.Blue,
                x => x.FontName = "Arial",
                x => x.FontFamilyNumbering = XlsxFontFamilyNumberingValues.Decorative,
                x => x.FontCharSet = XlsxFontCharSet.Arabic,
                x => x.FontScheme = XlsxFontScheme.Minor,
            };
            var cell = ws.Cell("A3");
            foreach (var modify in makeDifferentFont)
            {
                modify(cell.Style.Font);
                Assert.AreNotEqual(testFont, cell.Style.Font);
                cell = cell.CellRight();
            }
        }

        private static IEnumerable<object> FontApiSetters()
        {
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Bold, (font, value) => font.Bold = value, true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Bold, (font, value) => font.SetBold(value), true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Bold, (font, _) => font.SetBold(), true);

            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Italic, (font, value) => font.Italic = value, true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Italic, (font, value) => font.SetItalic(value), true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Italic, (font, _) => font.SetItalic(), true);

            var underlineValues = EnumPolyfill.GetValues<XlsxFontUnderlineValues>();
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Underline, (font, value) => font.Underline = value, underlineValues);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Underline, (font, value) => font.SetUnderline(value), underlineValues);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Underline, (font, _) => font.SetUnderline(), XlsxFontUnderlineValues.Single);

            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Strikethrough, (font, value) => font.Strikethrough = value, true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Strikethrough, (font, value) => font.SetStrikethrough(value), true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Strikethrough, (font, _) => font.SetStrikethrough(), true);

            var valignValues = EnumPolyfill.GetValues<XlsxFontVerticalTextAlignmentValues>();
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.VerticalAlignment, (font, value) => font.VerticalAlignment = value, valignValues);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.VerticalAlignment, (font, value) => font.SetVerticalAlignment(value), valignValues);

            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Shadow, (font, value) => font.Shadow = value, true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Shadow, (font, value) => font.SetShadow(value), true, false);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.Shadow, (font, _) => font.SetShadow(), true);

            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontSize, (font, value) => font.FontSize = value, 1, 15, 409.55);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontSize, (font, value) => font.SetFontSize(value), 1, 15, 409.55);

            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontColor, (font, value) => font.FontColor = value, XlsxColor.Black, XlsxColor.Red, XlsxColor.Automatic);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontColor, (font, value) => font.SetFontColor(value), XlsxColor.Black, XlsxColor.Red, XlsxColor.Automatic);

            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontName, (font, value) => font.FontName = value, "Calibri", "Arial", "Consolas");
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontName, (font, value) => font.SetFontName(value), "Calibri", "Arial", "Consolas");

            var familyValues = EnumPolyfill.GetValues<XlsxFontFamilyNumberingValues>();
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontFamilyNumbering, (font, value) => font.FontFamilyNumbering = value, familyValues);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontFamilyNumbering, (font, value) => font.SetFontFamilyNumbering(value), familyValues);

            var charsetValues = EnumPolyfill.GetValues<XlsxFontCharSet>();
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontCharSet, (font, value) => font.FontCharSet = value, charsetValues);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontCharSet, (font, value) => font.SetFontCharSet(value), charsetValues);

            var schemeValues = EnumPolyfill.GetValues<XlsxFontScheme>();
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontScheme, (font, value) => font.FontScheme = value, schemeValues);
            yield return FormatTestCase<IXlsxFont>.ForFont(font => font.FontScheme, (font, value) => font.SetFontScheme(value), schemeValues);
        }
    }
}
