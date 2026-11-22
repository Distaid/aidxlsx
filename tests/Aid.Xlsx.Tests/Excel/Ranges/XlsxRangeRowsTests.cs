using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel.Ranges;

internal class XlsxRangeRowsTests
{
    [Test]
    public void Style_sets_format_of_range_rows()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("B3:C7");
        var rangeRows = range.Rows("1,3-4");

        rangeRows.Style.Font.FontSize = 20;

        var expectedChangedCells = new[] { "B3", "C3", "B5", "C5", "B6", "C6" }.ToHashSet();
        foreach (var cell in range.Grow().Cells())
        {
            var address = cell.Address.ToString();
            var fontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(fontSize, cell.Style.Font.FontSize, 0, address);
        }
    }
}
