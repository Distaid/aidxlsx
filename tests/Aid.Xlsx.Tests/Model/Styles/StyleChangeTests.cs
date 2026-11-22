using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.Styles;

[TestFixture]
public class StyleChangeTests
{
    [Test]
    public void ChangeFontColorDoesNotAffectOtherProperties()
    {
        using (var wb = new XlsxWorkbook())
        {
            // Arrange
            var ws = wb.AddWorksheet("Sheet1");
            var a1 = ws.Cell("A1");
            var a2 = ws.Cell("A2");
            var b1 = ws.Cell("B1");
            var b2 = ws.Cell("B2");

            ws.Range("A1:B2").Value = "Test";

            a1.Style.Fill.BackgroundColor = XlsxColor.Red;
            a2.Style.Fill.BackgroundColor = XlsxColor.Green;
            b1.Style.Fill.BackgroundColor = XlsxColor.Blue;
            b2.Style.Fill.BackgroundColor = XlsxColor.Pink;

            a1.Style.Font.FontName = "Arial";
            a2.Style.Font.FontName = "Times New Roman";
            b1.Style.Font.FontName = "Calibri";
            b2.Style.Font.FontName = "Cambria";

            // Act
            ws.Range("A1:B2").Style.Font.FontColor = XlsxColor.PowderBlue;

            //Assert
            Assert.AreEqual(XlsxColor.Red, ws.Cell("A1").Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, ws.Cell("A2").Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Blue, ws.Cell("B1").Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Pink, ws.Cell("B2").Style.Fill.BackgroundColor);

            Assert.AreEqual("Arial", ws.Cell("A1").Style.Font.FontName);
            Assert.AreEqual("Times New Roman", ws.Cell("A2").Style.Font.FontName);
            Assert.AreEqual("Calibri", ws.Cell("B1").Style.Font.FontName);
            Assert.AreEqual("Cambria", ws.Cell("B2").Style.Font.FontName);

            Assert.AreEqual(XlsxColor.PowderBlue, ws.Cell("A1").Style.Font.FontColor);
            Assert.AreEqual(XlsxColor.PowderBlue, ws.Cell("A2").Style.Font.FontColor);
            Assert.AreEqual(XlsxColor.PowderBlue, ws.Cell("B1").Style.Font.FontColor);
            Assert.AreEqual(XlsxColor.PowderBlue, ws.Cell("B2").Style.Font.FontColor);
        }
    }

    [Test]
    public void ChangeStyleAlignment()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Style;

        style.Alignment.Horizontal = XlsxAlignmentHorizontalValues.Justify;

        Assert.AreEqual(XlsxAlignmentHorizontalValues.Justify, style.Alignment.Horizontal);
    }

    [Test]
    public void ChangeStyleBorder()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Style;

        style.Border.DiagonalBorder = XlsxBorderStyleValues.Double;

        Assert.AreEqual(XlsxBorderStyleValues.Double, style.Border.DiagonalBorder);
    }

    [Test]
    public void ChangeStyleFill()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Style;

        style.Fill.BackgroundColor = XlsxColor.Red;

        Assert.AreEqual(XlsxColor.Red, style.Fill.BackgroundColor);
    }

    [Test]
    public void ChangeStyleFont()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Style;

        style.Font.FontSize = 50;

        Assert.AreEqual(50, style.Font.FontSize);
    }

    [Test]
    public void ChangeStyleNumberFormat()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Style;

        style.NumberFormat.Format = "YYYY";

        Assert.AreEqual("YYYY", style.NumberFormat.Format);
    }

    [Test]
    public void ChangeStyleProtection()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Style;

        style.Protection.Hidden = true;

        Assert.AreEqual(true, style.Protection.Hidden);
    }

    [Test]
    public void ChangeAttachedStyleAlignment()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var a1 = ws.Cell("A1");

            a1.Style.Alignment.Horizontal = XlsxAlignmentHorizontalValues.Justify;

            Assert.AreEqual(XlsxAlignmentHorizontalValues.Justify, a1.Style.Alignment.Horizontal);
        }
    }
}
