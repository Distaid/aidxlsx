using NUnit.Framework;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.Styles;

[TestFixture]
public class StyleApplicationTests
{
    private static void SetupBorders(IXlsxRange range)
    {
        range.FirstRow().Cell(1).Style.Border.TopBorder = XlsxBorderStyleValues.None;
        range.FirstRow().Cell(2).Style.Border.TopBorder = XlsxBorderStyleValues.Thick;
        range.FirstRow().Cell(3).Style.Border.TopBorder = XlsxBorderStyleValues.Double;

        range.LastRow().Cell(1).Style.Border.BottomBorder = XlsxBorderStyleValues.None;
        range.LastRow().Cell(2).Style.Border.BottomBorder = XlsxBorderStyleValues.Thick;
        range.LastRow().Cell(3).Style.Border.BottomBorder = XlsxBorderStyleValues.Double;

        range.FirstColumn().Cell(1).Style.Border.LeftBorder = XlsxBorderStyleValues.None;
        range.FirstColumn().Cell(2).Style.Border.LeftBorder = XlsxBorderStyleValues.Thick;
        range.FirstColumn().Cell(3).Style.Border.LeftBorder = XlsxBorderStyleValues.Double;

        range.LastColumn().Cell(1).Style.Border.RightBorder = XlsxBorderStyleValues.None;
        range.LastColumn().Cell(2).Style.Border.RightBorder = XlsxBorderStyleValues.Thick;
        range.LastColumn().Cell(3).Style.Border.RightBorder = XlsxBorderStyleValues.Double;
    }

    [Test]
    public void InsideBorderTest()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.AddWorksheet("Sheet1");
        IXlsxRange range = ws.Range("B2:D4");

        SetupBorders(range);

        range.Style.Border.InsideBorder = XlsxBorderStyleValues.Thin;
        range.Style.Border.InsideBorderColor = XlsxColor.Red;

        IXlsxCell center = range.Cell(2, 2);

        Assert.AreEqual(XlsxColor.Red, center.Style.Border.TopBorderColor);
        Assert.AreEqual(XlsxColor.Red, center.Style.Border.BottomBorderColor);
        Assert.AreEqual(XlsxColor.Red, center.Style.Border.LeftBorderColor);
        Assert.AreEqual(XlsxColor.Red, center.Style.Border.RightBorderColor);

        Assert.AreEqual(XlsxBorderStyleValues.None, range.FirstRow().Cell(1).Style.Border.TopBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Thick, range.FirstRow().Cell(2).Style.Border.TopBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Double, range.FirstRow().Cell(3).Style.Border.TopBorder);

        Assert.AreEqual(XlsxBorderStyleValues.None, range.LastRow().Cell(1).Style.Border.BottomBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Thick, range.LastRow().Cell(2).Style.Border.BottomBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Double, range.LastRow().Cell(3).Style.Border.BottomBorder);

        Assert.AreEqual(XlsxBorderStyleValues.None, range.FirstColumn().Cell(1).Style.Border.LeftBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Thick, range.FirstColumn().Cell(2).Style.Border.LeftBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Double, range.FirstColumn().Cell(3).Style.Border.LeftBorder);

        Assert.AreEqual(XlsxBorderStyleValues.None, range.LastColumn().Cell(1).Style.Border.RightBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Thick, range.LastColumn().Cell(2).Style.Border.RightBorder);
        Assert.AreEqual(XlsxBorderStyleValues.Double, range.LastColumn().Cell(3).Style.Border.RightBorder);
    }

    [Test]
    public void ResolveThemeColors()
    {
        using (var wb = new XlsxWorkbook())
        {
            string color;
            color = wb.Theme.ResolveThemeColor(XlsxThemeColor.Accent1).Color.ToHex();
            Assert.AreEqual("FF4F81BD", color);

            color = wb.Theme.ResolveThemeColor(XlsxThemeColor.Background1).Color.ToHex();
            Assert.AreEqual("FFFFFFFF", color);
        }
    }

    [Theory]
    public void CanResolveAllThemeColors(XlsxThemeColor themeColor)
    {
        var theme = new XlsxWorkbook().Theme;
        var color = theme.ResolveThemeColor(themeColor);
        Assert.IsNotNull(color);
    }

    [Test]
    public void SetStyleViaRowReference()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Style
               .Font.SetFontSize(8)
               .Font.SetFontColor(XlsxColor.Green)
               .Font.SetBold(true);

            var row = ws.Row(1);
            ws.Cell(1, 1).Value = "Test";
            row.Cell(2).Value = "Test";
            row.Cells(3, 3).Value = "Test";

            foreach (var cell in ws.CellsUsed())
            {
                Assert.AreEqual(8, ws.Cell("A1").Style.Font.FontSize);
                Assert.AreEqual(XlsxColor.Green, ws.Cell("B1").Style.Font.FontColor);
                Assert.AreEqual(true, ws.Cell("C1").Style.Font.Bold);
            }
        }
    }
}
