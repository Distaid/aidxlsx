using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel.Ranges;

internal class XlsxRangesTests
{
    [Test]
    public void Style_sets_format_of_ranges()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var ranges = ws.Ranges("B3:C4,C7:D7");

        ranges.Style.Font.FontSize = 20;

        var expectedChangedCells = new[] { "B3", "C3", "B4", "C4", "C7", "D7" }.ToHashSet();
        foreach (var cell in ranges.Cells())
        {
            var address = cell.Address.ToString();
            var fontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(fontSize, cell.Style.Font.FontSize, 0, address);
        }
    }
}
