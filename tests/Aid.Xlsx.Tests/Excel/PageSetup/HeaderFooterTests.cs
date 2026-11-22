using NUnit.Framework;
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel;

[TestFixture]
public class HeaderFooterTests
{
    [Test]
    public void CanChangeWorksheetHeader()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.AddWorksheet("Sheet1");

        ws.PageSetup.Header.Center.AddText("Initial page header", XlsxHFOccurrence.EvenPages);

        var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        wb = new XlsxWorkbook(ms);
        ws = wb.Worksheets.First();

        ws.PageSetup.Header.Center.Clear();
        ws.PageSetup.Header.Center.AddText("Changed header", XlsxHFOccurrence.EvenPages);

        wb.SaveAs(ms, true);

        wb = new XlsxWorkbook(ms);
        ws = wb.Worksheets.First();

        var newHeader = ws.PageSetup.Header.Center.GetText(XlsxHFOccurrence.EvenPages);
        Assert.AreEqual("Changed header", newHeader);
    }

    [TestCase("")]
    [TestCase("&L&C&\"Arial\"&9 19-10-2017 \n&9&\"Arial\" &P    &N &R")] // upstream issue #563
    public void CanSetHeaderFooter(string s)
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            {
                var header = ws.PageSetup.Header as XlsxHeaderFooter;
                header.SetInnerText(XlsxHFOccurrence.AllPages, s);
            }
        }
    }
}
