using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel.CalcEngine;

public class CalcEngineCellEnumerationTests
{
    [Test]
    public void CanEnumerateCellsOverEmptySheet()
    {
        using (var wb = new XlsxWorkbook())
        {
            var sheet1 = wb.AddWorksheet("Sheet1");
            var sheet2 = wb.AddWorksheet("Sheet2");

            var cell = sheet1.FirstCell();
            cell.FormulaA1 = "=SUMIFS(Sheet2!B:B, Sheet2!C:C, 1)";

            Assert.AreEqual(0, cell.Value);
        }
    }
}
