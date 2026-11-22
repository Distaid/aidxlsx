using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.Hyperlinks;

[TestFixture]
public class HyperlinkTests
{
    [Test]
    public void TestHyperlinks()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Sheet2");

            var targetCell = ws2.Cell("A1");
            var targetRange = ws2.Range("A1", "B1");

            var linkCell1 = ws1.Cell("A1");
            linkCell1.Value = "Link to IXlsxCell";
            linkCell1.SetHyperlink(new XlsxHyperlink(targetCell));
            Assert.AreEqual("Sheet2!A1", linkCell1.GetHyperlink().InternalAddress);

            var linkRange1 = ws1.Cell("A2");
            linkRange1.Value = "Link to IXlsxRangeBase";
            linkRange1.SetHyperlink(new XlsxHyperlink(targetRange));
            Assert.AreEqual("Sheet2!A1:B1", linkRange1.GetHyperlink().InternalAddress);
        }
    }
}
