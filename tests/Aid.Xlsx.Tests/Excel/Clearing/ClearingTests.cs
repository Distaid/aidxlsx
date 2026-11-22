using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
using NUnit.Framework;

namespace Aid.Xlsx.Tests
{
    [TestFixture]
    public class ClearingTests
    {
        private static XlsxColor backgroundColor = XlsxColor.LightBlue;
        private static XlsxColor foregroundColor = XlsxColor.DarkBrown;

        private IXlsxWorkbook SetupWorkbook()
        {
            var wb = new XlsxWorkbook();
            IXlsxWorksheet ws = wb.Worksheets.Add("Sheet1");

            var c = ws.FirstCell()
                .SetValue("Hello world!");

            c.GetComment().AddText("Some comment");

            c.Style.Fill.BackgroundColor = backgroundColor;
            c.Style.Font.FontColor = foregroundColor;
            c.CreateDataValidation().Custom("B1");

            ////

            c = ws.FirstCell()
                .CellBelow()
                .SetFormulaA1("=LEFT(A1,5)");

            c.GetComment().AddText("Another comment");

            c.Style.Fill.BackgroundColor = backgroundColor;
            c.Style.Font.FontColor = foregroundColor;

            ////

            c = ws.FirstCell()
                .CellBelow(2)
                .SetValue(new DateTime(2018, 1, 15));

            c.GetComment().AddText("A date");

            c.Style.Fill.BackgroundColor = backgroundColor;
            c.Style.Font.FontColor = foregroundColor;

            ws.Column(1)
                .AddConditionalFormat().WhenStartsWith("Hell")
                .Fill.SetBackgroundColor(XlsxColor.Red)
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XlsxColor.Blue)
                .Font.SetBold();

            Assert.AreEqual(XlsxDataType.Text, ws.Cell("A1").Value.Type);
            Assert.AreEqual(XlsxDataType.Text, ws.Cell("A2").Value.Type);
            Assert.AreEqual(XlsxDataType.DateTime, ws.Cell("A3").Value.Type);

            Assert.AreEqual(false, ws.Cell("A1").HasFormula);
            Assert.AreEqual(true, ws.Cell("A2").HasFormula);
            Assert.AreEqual(false, ws.Cell("A1").HasFormula);

            foreach (var cell in ws.Range("A1:A3").Cells())
            {
                Assert.AreEqual(backgroundColor, cell.Style.Fill.BackgroundColor);
                Assert.AreEqual(foregroundColor, cell.Style.Font.FontColor);
                Assert.IsTrue(ws.ConditionalFormats.Any());
                Assert.IsTrue(cell.HasComment);
            }

            Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);

            return wb;
        }

        [Test]
        public void WorksheetClearAll()
        {
            using (var wb = SetupWorkbook())
            {
                var ws = wb.Worksheets.First();

                ws.Clear(XlsxClearOptions.All);

                foreach (var c in ws.Range("A1:A10").Cells())
                {
                    Assert.IsTrue(c.IsEmpty());
                    Assert.AreEqual(XlsxDataType.Blank, c.DataType);
                    Assert.AreEqual(ws.Style.Fill.BackgroundColor, c.Style.Fill.BackgroundColor);
                    Assert.AreEqual(ws.Style.Font.FontColor, c.Style.Font.FontColor);
                    Assert.IsFalse(ws.ConditionalFormats.Any());
                    Assert.IsFalse(c.HasComment);
                    Assert.AreEqual(String.Empty, c.GetDataValidation().Value);
                }
            }
        }

        [Test]
        public void WorksheetClearContents()
        {
            using (var wb = SetupWorkbook())
            {
                var ws = wb.Worksheets.First();

                ws.Clear(XlsxClearOptions.Contents);

                foreach (var c in ws.Range("A1:A3").Cells())
                {
                    Assert.AreEqual(XlsxDataType.Blank, ws.Cell("A1").DataType);
                    Assert.IsTrue(c.IsEmpty(XlsxCellsUsedOptions.Contents));

                    Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                    Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                    Assert.IsTrue(ws.ConditionalFormats.Any());
                    Assert.IsTrue(c.HasComment);
                }

                Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
            }
        }

        [Test]
        public void WorksheetClearNormalFormats()
        {
            using (var wb = SetupWorkbook())
            {
                var ws = wb.Worksheets.First();

                ws.Clear(XlsxClearOptions.NormalFormats);

                foreach (var c in ws.Range("A1:A3").Cells())
                {
                    Assert.IsFalse(c.IsEmpty());
                    Assert.AreEqual(ws.Style.Fill.BackgroundColor, c.Style.Fill.BackgroundColor);
                    Assert.AreEqual(ws.Style.Font.FontColor, c.Style.Font.FontColor);
                    Assert.IsTrue(ws.ConditionalFormats.Any());
                    Assert.IsTrue(c.HasComment);
                }

                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A1").DataType);
                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A2").DataType);
                Assert.AreEqual(XlsxDataType.DateTime, ws.Cell("A3").DataType);

                Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
            }
        }

        [Test]
        public void WorksheetClearConditionalFormats()
        {
            using (var wb = SetupWorkbook())
            {
                var ws = wb.Worksheets.First();

                ws.Clear(XlsxClearOptions.ConditionalFormats);

                foreach (var c in ws.Range("A1:A3").Cells())
                {
                    Assert.IsFalse(c.IsEmpty());
                    Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                    Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                    Assert.IsFalse(ws.ConditionalFormats.Any());
                    Assert.IsTrue(c.HasComment);
                }

                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A1").DataType);
                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A2").DataType);
                Assert.AreEqual(XlsxDataType.DateTime, ws.Cell("A3").DataType);

                Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
            }
        }

        [Test]
        public void WorksheetClearComments()
        {
            using (var wb = SetupWorkbook())
            {
                var ws = wb.Worksheets.First();

                ws.Clear(XlsxClearOptions.Comments);

                foreach (var c in ws.Range("A1:A3").Cells())
                {
                    Assert.IsFalse(c.IsEmpty());
                    Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                    Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                    Assert.IsTrue(ws.ConditionalFormats.Any());
                    Assert.IsFalse(c.HasComment);
                }

                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A1").DataType);
                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A2").DataType);
                Assert.AreEqual(XlsxDataType.DateTime, ws.Cell("A3").DataType);

                Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
            }
        }

        [Test]
        public void WorksheetClearDataValidation()
        {
            using (var wb = SetupWorkbook())
            {
                var ws = wb.Worksheets.First();

                ws.Clear(XlsxClearOptions.DataValidation);

                foreach (var c in ws.Range("A1:A3").Cells())
                {
                    Assert.IsFalse(c.IsEmpty());
                    Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                    Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                    Assert.IsTrue(ws.ConditionalFormats.Any());
                    Assert.IsTrue(c.HasComment);
                }

                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A1").DataType);
                Assert.AreEqual(XlsxDataType.Text, ws.Cell("A2").DataType);
                Assert.AreEqual(XlsxDataType.DateTime, ws.Cell("A3").DataType);

                Assert.AreEqual(string.Empty, ws.Cell("A1").GetDataValidation().Value);
            }
        }

        [Test]
        public void DeleteClearedCellValue()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = SetupWorkbook())
                {
                    var ws = wb.Worksheets.First();
                    Assert.AreEqual("Hello world!", ws.Cell("A1").GetText());
                    Assert.AreEqual(new DateTime(2018, 1, 15), ws.Cell("A3").GetDateTime());

                    wb.SaveAs(ms);
                }

                using (var wb = new XlsxWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    ws.Clear(XlsxClearOptions.Contents);
                    Assert.AreEqual(Blank.Value, ws.Cell("A1").Value);
                    Assert.Throws<InvalidCastException>(() => ws.Cell("A3").GetDateTime());

                    wb.Save();
                }

                using (var wb = new XlsxWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    Assert.AreEqual(Blank.Value, ws.Cell("A1").Value);
                    Assert.Throws<InvalidCastException>(() => ws.Cell("A3").GetDateTime());
                }
            }
        }

        [TestCase(XlsxClearOptions.All, 2)]
        [TestCase(XlsxClearOptions.AllContents, 4)]
        [TestCase(XlsxClearOptions.AllFormats, 4)]
        [TestCase(XlsxClearOptions.Contents, 4)]
        [TestCase(XlsxClearOptions.MergedRanges, 2)]
        public void CanClearMergedRanges(XlsxClearOptions options, int expectedCount)
        {
            using (var wb = new XlsxWorkbook())
            {
                var ws = wb.AddWorksheet("Test");

                ws.Range("A1:C3").Merge();
                ws.Range("A4:B6").Merge();
                ws.Range("D1:F3").Merge();
                ws.Range("E4:F6").Merge();

                ws.Range("C1:D6").Clear(options);

                Assert.AreEqual(expectedCount, ws.MergedRanges.Count);
            }
        }
    }
}
