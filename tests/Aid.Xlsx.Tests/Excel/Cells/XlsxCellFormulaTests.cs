using Aid.Xlsx.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.Cells
{
    [TestFixture]
    public class XlsxCellFormulaTests
    {
        [Test]
        public void CellFormulaIsStrippedOfEqualSign()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 1).FormulaA1 = "=B1";
            Assert.AreEqual("B1", ws.Cell(1, 1).FormulaA1);
        }

        [Test]
        public void DataTable_MaintainProperties()
        {
            TestHelper.LoadSaveAndCompare(
                @"Other\Formulas\DataTableFormula-Excel-Input.xlsx",
                @"Other\Formulas\DataTableFormula-Output.xlsx");
        }
    }
}
