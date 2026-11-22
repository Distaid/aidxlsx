using NUnit.Framework;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Rendering;

namespace Aid.Xlsx.Tests.Graphics;

[TestFixture]
public class FontTests
{
    private readonly IXlsxGraphicEngine _engine = DefaultGraphicEngine.Instance.Value;

    [TestCase]
    public void CalculatedTextWidth()
    {
        var textFont = new DummyFont("Calibri", 20);
        var textWidthPt = _engine.GetTextWidth("Lorem ipsum dolor sit amet", textFont, 96);
        Assert.That(textWidthPt, Is.EqualTo(300));
    }

    [TestCase]
    public void CalculatedTextHeight()
    {
        // Calibri is measured through the embedded CarlitoBare font on every platform, so the value
        // is platform independent (Carlito's ascender is 1950/2048 and descender -550/2048).
        var textFont = new DummyFont("Calibri", 300);
        var textHeightPx = _engine.GetTextHeight(textFont, 96);
        Assert.That(textHeightPx, Is.EqualTo(595.703125));
    }

    [TestCase]
    public void GetMaxDigitWidth()
    {
        var textFont = new DummyFont("Calibri", 11);
        var textWidthPx = _engine.GetMaxDigitWidth(textFont, 96);
        Assert.That(textWidthPx, Is.EqualTo(7.43359375d)); // Calibri,11 has a max digit width of 7 per spec 18.3.1.13
    }

    [TestCase]
    public void DescentIsPositive()
    {
        // See CalculatedTextHeight: Carlito descent is -550/2048 of the em square.
        var textFont = new DummyFont("Calibri", 11);
        var textWidthPt = _engine.GetDescent(textFont, 96);
        Assert.That(textWidthPt, Is.EqualTo(3.938802083333333d));
    }

    [TestCase]
    public void NonExistentFontUsesFallback()
    {
        var nonExistentFont = new DummyFont("NonExistentFont", 100);
        var fallbackFont = new DummyFont("Microsoft Sans Serif", 100);

        var nonExistentFontWidth = _engine.GetTextWidth("ABCDEF text", nonExistentFont, 96);
        var fallbackFontWidth = _engine.GetTextWidth("ABCDEF text", fallbackFont, 96);
        Assert.That(nonExistentFontWidth, Is.EqualTo(fallbackFontWidth));

        var nonExistentFontHeight = _engine.GetTextHeight(nonExistentFont, 96);
        var fallbackFontHeight = _engine.GetTextHeight(fallbackFont, 96);
        Assert.That(nonExistentFontHeight, Is.EqualTo(fallbackFontHeight));
    }

    [Test]
    public void UseEmbeddedFontWhenFallbackFontIsNotPresent()
    {
        var nonExistentFont = new DummyFont("SomeNonExistentFont", 11);
        var engine = new DefaultGraphicEngine("NonExistentFallbackFont");
        Span<int> text = stackalloc int[1] { '8' };

        var box = engine.GetGlyphBox(text, nonExistentFont, new Dpi(96, 96));

        // Max digit width of CarlitoBare is 7, unlike MS Sans Serif which is 8
        Assert.AreEqual(7, box.AdvanceWidth);
    }

    [TestCase]
    public void CanSpecifyFallbackFontWithoutFileSystem()
    {
        using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = DefaultGraphicEngine.CreateOnlyWithFonts(fallbackFontStream);

        var nonExistentFont = new DummyFont("Nonexistent Font", 20);
        var widthOfLetterA = engine.GetTextWidth("A", nonExistentFont, 120);

        const double expectedWidthOfLetterA = 31.25d;
        Assert.AreEqual(expectedWidthOfLetterA, widthOfLetterA, 0.0001);
    }

    [TestCase]
    public void CanSpecifyExtraFontsAsStreamsWithoutFileSystem()
    {
        using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        var engine = DefaultGraphicEngine.CreateOnlyWithFonts(fallbackFontStream, fontBStream);

        var widthOfLetterB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);

        const double expectedWidthOfLetterB = 25d;
        Assert.AreEqual(expectedWidthOfLetterB, widthOfLetterB, 0.0001);
    }

    [TestCase]
    public void Issue_1916_CanMeasureSpecificArabicText()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = @"اصين";
        ws.Column(1).AdjustToContents();
    }

    private class DummyFont : IXlsxFontBase
    {
        public DummyFont(string name, double size)
        {
            FontName = name;
            FontSize = size;
        }

        public string FontName { get; set; }

        public double FontSize { get; set; }

        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public bool Strikethrough { get; set; }

        public XlsxFontUnderlineValues Underline { get; set; } = XlsxFontUnderlineValues.None;

        public XlsxFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

        public bool Shadow { get; set; }

        public XlsxColor FontColor { get; set; } = XlsxColor.Black;

        public XlsxFontFamilyNumberingValues FontFamilyNumbering { get; set; } = XlsxFontFamilyNumberingValues.NotApplicable;

        public XlsxFontCharSet FontCharSet { get; set; } = XlsxFontCharSet.Default;

        public XlsxFontScheme FontScheme { get; set; }
    }
}
