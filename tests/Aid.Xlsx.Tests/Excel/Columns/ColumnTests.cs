using Aid.Xlsx.Cells;
using Aid.Xlsx.Columns;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel
{
    [TestFixture]
    public class ColumnTests
    {
        [Test]
        public void ColumnUsed()
        {
            var wb = new XlsxWorkbook();
            IXlsxWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(2, 1).SetValue("Test");
            ws.Cell(3, 1).SetValue("Test");

            IXlsxRangeColumn fromColumn = ws.Column(1).ColumnUsed();
            Assert.AreEqual("A2:A3", fromColumn.RangeAddress.ToStringRelative());

            IXlsxRangeColumn fromRange = ws.Range("A1:A5").FirstColumn().ColumnUsed();
            Assert.AreEqual("A2:A3", fromRange.RangeAddress.ToStringRelative());
        }

        [Test]
        public void ColumnsUsedIsFast()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            var columnsUsed = ws.Row(1).AsRange().ColumnsUsed();
            Assert.AreEqual(1, columnsUsed.Count());
        }

        [Test]
        public void CopyColumn()
        {
            var wb = new XlsxWorkbook();
            IXlsxWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Test").Style.Font.SetBold();
            ws.FirstColumn().CopyTo(ws.Column(2));

            Assert.IsTrue(ws.Cell("B1").Style.Font.Bold);
        }

        [Test]
        public void InsertingColumnsBefore1()
        {
            var wb = new XlsxWorkbook();
            IXlsxWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XlsxColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XlsxColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XlsxColor.Green);

            IXlsxColumn column1 = ws.Column(1);
            IXlsxColumn column2 = ws.Column(2);
            IXlsxColumn column3 = ws.Column(3);

            IXlsxColumn columnIns = ws.Column(1).InsertColumnsBefore(1).First();

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, ws.Column(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Yellow, ws.Column(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, ws.Column(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, ws.Column(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Column(3).Cell(2).GetText());

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, columnIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, columnIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(ws.Style.Fill.BackgroundColor, columnIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, column1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Yellow, column2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, column2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, column2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, column3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", column2.Cell(2).GetText());
        }

        [Test]
        public void InsertingColumnsBefore2()
        {
            var wb = new XlsxWorkbook();
            IXlsxWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XlsxColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XlsxColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XlsxColor.Green);

            IXlsxColumn column1 = ws.Column(1);
            IXlsxColumn column2 = ws.Column(2);
            IXlsxColumn column3 = ws.Column(3);

            IXlsxColumn columnIns = ws.Column(2).InsertColumnsBefore(1).First();

            Assert.AreEqual(XlsxColor.Red, ws.Column(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, ws.Column(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Yellow, ws.Column(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, ws.Column(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, ws.Column(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Column(3).Cell(2).GetText());

            Assert.AreEqual(XlsxColor.Red, columnIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, columnIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, columnIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, column1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Yellow, column2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, column2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, column2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, column3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", column2.Cell(2).GetText());
        }

        [Test]
        public void InsertingColumnsBefore3()
        {
            var wb = new XlsxWorkbook();
            IXlsxWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Columns("1,3").Style.Fill.SetBackgroundColor(XlsxColor.Red);
            ws.Column(2).Style.Fill.SetBackgroundColor(XlsxColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XlsxColor.Green);

            IXlsxColumn column1 = ws.Column(1);
            IXlsxColumn column2 = ws.Column(2);
            IXlsxColumn column3 = ws.Column(3);

            IXlsxColumn columnIns = ws.Column(3).InsertColumnsBefore(1).First();

            Assert.AreEqual(XlsxColor.Red, ws.Column(1).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(1).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(1).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Yellow, ws.Column(2).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, ws.Column(2).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, ws.Column(2).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Yellow, ws.Column(3).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, ws.Column(3).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, ws.Column(3).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, ws.Column(4).Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", ws.Column(2).Cell(2).GetText());

            Assert.AreEqual(XlsxColor.Yellow, columnIns.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, columnIns.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, columnIns.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, column1.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column1.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column1.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Yellow, column2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Green, column2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Yellow, column2.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual(XlsxColor.Red, column3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XlsxColor.Red, column3.Cell(3).Style.Fill.BackgroundColor);

            Assert.AreEqual("X", column2.Cell(2).GetText());
        }

        [Test]
        public void NoColumnsUsed()
        {
            var wb = new XlsxWorkbook();
            IXlsxWorksheet ws = wb.Worksheets.Add("Sheet1");
            Int32 count = 0;

            foreach (IXlsxColumn row in ws.ColumnsUsed())
                count++;

            foreach (IXlsxRangeColumn row in ws.Range("A1:C3").ColumnsUsed())
                count++;

            Assert.AreEqual(0, count);
        }

        [Test]
        public void UngroupFromAll()
        {
            IXlsxWorksheet ws = new XlsxWorkbook().AddWorksheet("Sheet1");
            ws.Columns(1, 2).Group();
            ws.Columns(1, 2).Ungroup(true);
        }

        [Test]
        public void LastColumnUsed()
        {
            IXlsxWorksheet ws = new XlsxWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "A1";
            ws.Cell("B1").Value = "B1";
            ws.Cell("A2").Value = "A2";
            var lastCoUsed = ws.LastColumnUsed().ColumnNumber();
            Assert.AreEqual(2, lastCoUsed);
        }

        [Test]
        public void NegativeColumnNumberIsInvalid()
        {
            var ws = new XlsxWorkbook().AddWorksheet("Sheet1") as XlsxWorksheet;

            var column = new XlsxColumn(ws, -1);

            Assert.IsFalse(column.RangeAddress.IsValid);
        }

        [Test]
        public void AssignWorksheetColumnWidthWhenAllColumnsChanged()
        {
            var ws = new XlsxWorkbook().AddWorksheet();
            var columns = ws.Columns();

            columns.Width = 100;

            Assert.AreEqual(100, ws.Column("G").Width, XlsxHelper.Epsilon);
            Assert.AreEqual(100, ws.ColumnWidth, XlsxHelper.Epsilon);
        }

        [Test]
        public void PreserveWorksheetColumnWidthWhenNotAllColumnsChanged()
        {
            var ws = new XlsxWorkbook().AddWorksheet();
            var defaultColumnWidth = ws.ColumnWidth;
            var columns = ws.Columns(1, XlsxHelper.MaxColumnNumber);

            columns.Width = 100;

            Assert.AreEqual(100, ws.Column("G").Width, XlsxHelper.Epsilon);
            Assert.AreEqual(defaultColumnWidth, ws.ColumnWidth, XlsxHelper.Epsilon);
        }

        [Test]
        public void PreserveWorksheetColumnWidthWhenUsedColumnsChanged()
        {
            var ws = new XlsxWorkbook().AddWorksheet();
            ws.Cells("A1:E5").Value = "Not empty";
            var defaultColumnWidth = ws.ColumnWidth;
            var columns = ws.ColumnsUsed(XlsxCellsUsedOptions.Contents);

            columns.Width = 100;

            Assert.AreEqual(100, ws.Column("C").Width, XlsxHelper.Epsilon);
            Assert.AreEqual(defaultColumnWidth, ws.Column("G").Width, XlsxHelper.Epsilon);
            Assert.AreEqual(defaultColumnWidth, ws.ColumnWidth, XlsxHelper.Epsilon);
        }
    }
}
