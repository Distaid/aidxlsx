using System.Globalization;
using System.Xml.Linq;
using NUnit.Framework;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.SpreadsheetML.Styles;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.Tests.Model.Styles;

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

    [TestCase(OpenXmlConst.Main2006SsNs)]
    [TestCase(OpenXmlConst.X14Main2009SsNs)]
    public void SpreadsheetColorsRoundTripThroughXml(string namespaceName)
    {
        var name = XName.Get("color", namespaceName);
        var red = SpreadsheetColorXml.Write(name, XlsxColor.Red);
        var indexed = SpreadsheetColorXml.Write(name, XlsxColor.FromIndex(20));
        var theme = SpreadsheetColorXml.Write(name, XlsxColor.FromTheme(XlsxThemeColor.Accent1));
        var tint = SpreadsheetColorXml.Write(name, XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.4));

        Assert.That(red.Attributes().Select(attribute => attribute.ToString()), Is.EqualTo(new[] { "rgb=\"FFFF0000\"" }));
        Assert.That(indexed.Attributes().Select(attribute => attribute.ToString()), Is.EqualTo(new[] { "indexed=\"20\"" }));
        Assert.That(theme.Attributes().Select(attribute => attribute.ToString()), Is.EqualTo(new[] { "theme=\"4\"" }));
        Assert.That(tint.Attributes().Select(attribute => attribute.ToString()), Is.EqualTo(new[] { "theme=\"5\"", "tint=\"0.4\"" }));
        Assert.That(SpreadsheetColorXml.Read(red), Is.EqualTo(XlsxColor.Red));
        Assert.That(SpreadsheetColorXml.Read(indexed), Is.EqualTo(XlsxColor.FromIndex(20)));
        Assert.That(SpreadsheetColorXml.Read(theme), Is.EqualTo(XlsxColor.FromTheme(XlsxThemeColor.Accent1)));
        Assert.That(SpreadsheetColorXml.Read(tint), Is.EqualTo(XlsxColor.FromTheme(XlsxThemeColor.Accent2, 0.4)));
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
