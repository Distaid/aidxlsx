using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Tests.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.Styles
{
    public class BorderTests
    {
        private const int None = 0x0;
        private const int Left = 0x1;
        private const int Right = 0x2;
        private const int Top = 0x4;
        private const int Bottom = 0x8;
        private const int All = Left | Top | Right | Bottom;

        [Test]
        public void InsideBorders_preserve_outside_borders()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            var halfAccent1Color = XlsxColor.FromTheme(XlsxThemeColor.Accent1, 0.5);
            ws.Cells("B2:C2").Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(halfAccent1Color);

            // Check pre-conditions
            AssertCellBorder(ws, "B2", Left | Right | Top | Bottom, XlsxBorderStyleValues.Thin, halfAccent1Color);
            AssertCellBorder(ws, "C2", Left | Right | Top | Bottom, XlsxBorderStyleValues.Thin, halfAccent1Color);

            ws.Range("B2:C2").Style.Border.SetInsideBorder(XlsxBorderStyleValues.None);

            AssertCellBorder(ws, "B2", Left | Top | Bottom, XlsxBorderStyleValues.Thin);
            AssertCellBorder(ws, "C2", Right | Top | Bottom, XlsxBorderStyleValues.Thin);
            Assert.AreEqual(XlsxThemeColor.Accent1, ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor);
            Assert.AreEqual(XlsxThemeColor.Accent1, ws.Cell("C2").Style.Border.RightBorderColor.ThemeColor);
        }

        [Test]
        public void InsideBorder_sets_border_that_is_not_on_edge()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D4").Style
                .Border.SetInsideBorderColor(XlsxColor.Red)
                .Border.SetInsideBorder(XlsxBorderStyleValues.Thick);

            AssertCellBorder(ws, "B2", Bottom | Right);
            AssertCellBorder(ws, "C2", Left | Bottom | Right);
            AssertCellBorder(ws, "D2", Bottom | Left);
            AssertCellBorder(ws, "B3", Top | Right | Bottom);
            AssertCellBorder(ws, "C3", Left | Top | Right | Bottom);
            AssertCellBorder(ws, "D3", Top | Left | Bottom);
            AssertCellBorder(ws, "B4", Top | Right);
            AssertCellBorder(ws, "C4", Left | Top | Right);
            AssertCellBorder(ws, "D4", Left | Top);
        }

        [Test]
        public void OutsideBorder_sets_only_borders_on_edge()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D4").Style
                .Border.SetOutsideBorderColor(XlsxColor.Red)
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thick);

            AssertCellBorder(ws, "B2", Top | Left);
            AssertCellBorder(ws, "C2", Top);
            AssertCellBorder(ws, "D2", Top | Right);
            AssertCellBorder(ws, "B3", Left);
            AssertCellBorder(ws, "C3", None);
            AssertCellBorder(ws, "D3", Right);
            AssertCellBorder(ws, "B4", Bottom | Left);
            AssertCellBorder(ws, "C4", Bottom);
            AssertCellBorder(ws, "D4", Bottom | Right);
        }

        [Test]
        public void OutsideBorder_used_for_cells_will_make_outside_border_around_each_cell()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cells("B2:C3").Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XlsxColor.Red);

            AssertCellBorder(ws, "B2", All, XlsxBorderStyleValues.Thick, XlsxColor.Red);
            AssertCellBorder(ws, "C2", All, XlsxBorderStyleValues.Thick, XlsxColor.Red);
            AssertCellBorder(ws, "B3", All, XlsxBorderStyleValues.Thick, XlsxColor.Red);
            AssertCellBorder(ws, "C3", All, XlsxBorderStyleValues.Thick, XlsxColor.Red);
        }

        [Test]
        public void InsideBorder_used_for_cells_is_ignored_because_individual_cells_have_no_inside()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cells("B2:C3").Style
                .Border.SetInsideBorder(XlsxBorderStyleValues.Thick)
                .Border.SetInsideBorderColor(XlsxColor.Red);

            AssertCellBorder(ws, "B2", None, XlsxBorderStyleValues.None);
            AssertCellBorder(ws, "C2", None, XlsxBorderStyleValues.None);
            AssertCellBorder(ws, "B3", None, XlsxBorderStyleValues.None);
            AssertCellBorder(ws, "C3", None, XlsxBorderStyleValues.None);
        }

        [Test]
#if NETFRAMEWORK
        [Timeout(100)]
#else
        [MaxTime(100)]
#endif
        public void OutsideBorder_for_column()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between column and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XlsxColor.Blue);

            // Set border for a whole column
            ws.Column("B").Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XlsxColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.TopBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.TopBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.BottomBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.BottomBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.LeftBorder);
            Assert.AreEqual(XlsxColor.Red, b2.LeftBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.RightBorder);
            Assert.AreEqual(XlsxColor.Red, b2.RightBorderColor);
            AssertCellBorder(ws, "B1", Left | Right, XlsxBorderStyleValues.Thick, XlsxColor.Red);
            AssertCellBorder(ws, "B3", Left | Right, XlsxBorderStyleValues.Thick, XlsxColor.Red);
        }

        [Test]
#if NETFRAMEWORK
        [Timeout(100)]
#else
        [MaxTime(100)]
#endif
        public void InsideBorder_for_one_column()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between column and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XlsxColor.Blue);

            // Set border for a whole column
            ws.Column("B").Style
                .Border.SetInsideBorder(XlsxBorderStyleValues.Thick)
                .Border.SetInsideBorderColor(XlsxColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.TopBorder);
            Assert.AreEqual(XlsxColor.Red, b2.TopBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.BottomBorder);
            Assert.AreEqual(XlsxColor.Red, b2.BottomBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.LeftBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.LeftBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.RightBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.RightBorderColor);
            AssertCellBorder(ws, "B1", Top | Bottom, XlsxBorderStyleValues.Thick, XlsxColor.Red);
            AssertCellBorder(ws, "B3", Top | Bottom, XlsxBorderStyleValues.Thick, XlsxColor.Red);
        }

        [Test]
        public void InsideBorder_for_multicolumn_colspans()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // Reordered B-C,E-G - It can be in any order, duplicates are allowed in column specification string
            ws.Columns("E,G,C,F,B,E,C").Style.Border.SetInsideBorder(XlsxBorderStyleValues.Thin);

            AssertInsideBorderLeftColumn('B');
            AssertInsideBorderRightColumn('C');

            AssertInsideBorderLeftColumn('E');
            AssertInsideBorderCenterColumns('F');
            AssertInsideBorderRightColumn('G');
            return;

            void AssertInsideBorderLeftColumn(char column) => AssertColumn(column, Top | Right | Bottom);
            void AssertInsideBorderCenterColumns(char column) => AssertColumn(column, All);
            void AssertInsideBorderRightColumn(char column) => AssertColumn(column, Left | Top | Bottom);

            void AssertColumn(char column, int sides)
            {
                AssertCellBorder(ws, $"{column}1", sides, XlsxBorderStyleValues.Thin);
                AssertCellBorder(ws, $"{column}2", sides, XlsxBorderStyleValues.Thin);
                AssertCellBorder(ws, $"{column}10", sides, XlsxBorderStyleValues.Thin);
            }
        }

        [Test]
        public void OutsideBorder_for_rows()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between row and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XlsxColor.Blue);

            // Set border for a whole row
            ws.Row(2).Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XlsxColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.LeftBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.LeftBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.RightBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.RightBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.TopBorder);
            Assert.AreEqual(XlsxColor.Red, b2.TopBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.BottomBorder);
            Assert.AreEqual(XlsxColor.Red, b2.BottomBorderColor);

            AssertCellBorder(ws, "A2", Top | Bottom, XlsxBorderStyleValues.Thick, XlsxColor.Red);
            AssertCellBorder(ws, "C2", Top | Bottom, XlsxBorderStyleValues.Thick, XlsxColor.Red);
        }

        [Test]
        public void InsideBorder_for_one_row()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between row and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XlsxBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XlsxColor.Blue);

            // Set border for a whole row
            ws.Row(2).Style
                .Border.SetInsideBorder(XlsxBorderStyleValues.Thick)
                .Border.SetInsideBorderColor(XlsxColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.LeftBorder);
            Assert.AreEqual(XlsxColor.Red, b2.LeftBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thick, b2.RightBorder);
            Assert.AreEqual(XlsxColor.Red, b2.RightBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.TopBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.TopBorderColor);
            Assert.AreEqual(XlsxBorderStyleValues.Thin, b2.BottomBorder);
            Assert.AreEqual(XlsxColor.Blue, b2.BottomBorderColor);

            AssertCellBorder(ws, "A2", Left | Right, XlsxBorderStyleValues.Thick, XlsxColor.Red);
            AssertCellBorder(ws, "C2", Left | Right, XlsxBorderStyleValues.Thick, XlsxColor.Red);
        }

        [Test]
        public void InsideBorder_for_multirow_rowspans()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // Reordered 2-3,5-7 - It can be in any order, duplicates are allowed in row specification string
            ws.Rows("6,2,7,6,3,5,7").Style.Border.SetInsideBorder(XlsxBorderStyleValues.Thin);

            AssertInsideBorderTopRow(2);
            AssertInsideBorderBottomRow(3);

            AssertInsideBorderTopRow(5);
            AssertInsideBorderCenterRow(6);
            AssertInsideBorderBottomRow(7);
            return;

            void AssertInsideBorderTopRow(int row) => AssertRow(row, Left | Bottom | Right);
            void AssertInsideBorderCenterRow(int row) => AssertRow(row, All);
            void AssertInsideBorderBottomRow(int row) => AssertRow(row, Left | Top | Right);

            void AssertRow(int row, int sides)
            {
                AssertCellBorder(ws, $"A{row}", sides, XlsxBorderStyleValues.Thin);
                AssertCellBorder(ws, $"B{row}", sides, XlsxBorderStyleValues.Thin);
                AssertCellBorder(ws, $"Z{row}", sides, XlsxBorderStyleValues.Thin);
            }
        }

        [Test]
        [TestCaseSource(nameof(BorderApiSetters))]
        public void Border_property_can_be_individually_set(FormatTestCase<IXlsxBorder> testCase)
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("B2");

            foreach (var testValue in testCase.Values)
            {
                testCase.SetPropertyValue(cell.Style.Border, testValue);
                var setValue = testCase.GetPropertyValue(cell.Style.Border);
                Assert.AreEqual(testValue, setValue);
                cell = cell.CellRight();
            }
        }

        [TestCaseSource(nameof(BorderColorSetters))]
        public void Color_is_set_only_when_border_is_visible(Func<IXlsxBorder, XlsxColor> getColor, Action<IXlsxBorder, XlsxColor> setColor, Action<IXlsxBorder, XlsxBorderStyleValues> setStyle)
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            var cell = ws.Cell("A1");
            var defaultBorderColor = getColor(cell.Style.Border);
            var color = XlsxColor.Red;

            // Try and fail to set color, when border style is None.
            setStyle(cell.Style.Border, XlsxBorderStyleValues.None);
            setColor(cell.Style.Border, color);
            Assert.AreEqual(defaultBorderColor, getColor(cell.Style.Border));

            // Set color, when border style is visible.
            setStyle(cell.Style.Border, XlsxBorderStyleValues.Thin);
            setColor(cell.Style.Border, color);
            Assert.AreEqual(color, getColor(cell.Style.Border));
        }

        [TestCaseSource(nameof(BorderColorSetters))]
        public void Making_border_hidden_resets_the_color_to_default(Func<IXlsxBorder, XlsxColor> getColor, Action<IXlsxBorder, XlsxColor> setColor, Action<IXlsxBorder, XlsxBorderStyleValues> setStyle)
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            var cell = ws.Cell("A1");
            var defaultBorderColor = getColor(cell.Style.Border);
            var color = XlsxColor.Red;

            // Set the color of visible border
            setStyle(cell.Style.Border, XlsxBorderStyleValues.Thin);
            setColor(cell.Style.Border, color);
            Assert.AreEqual(color, getColor(cell.Style.Border));

            // When the border is hidden, the color is reset to default border color
            setStyle(cell.Style.Border, XlsxBorderStyleValues.None);
            Assert.AreEqual(defaultBorderColor, getColor(cell.Style.Border));
        }

        private static void AssertCellBorder(IXlsxWorksheet ws, string cell, int sides, XlsxBorderStyleValues style = XlsxBorderStyleValues.Thick, XlsxColor color = null)
        {
            var border = ws.Cell(cell).Style.Border;
            Assert.AreEqual((sides & Left) != 0 ? style : XlsxBorderStyleValues.None, border.LeftBorder);
            Assert.AreEqual((sides & Right) != 0 ? style : XlsxBorderStyleValues.None, border.RightBorder);
            Assert.AreEqual((sides & Top) != 0 ? style : XlsxBorderStyleValues.None, border.TopBorder);
            Assert.AreEqual((sides & Bottom) != 0 ? style : XlsxBorderStyleValues.None, border.BottomBorder);

            if (color is not null)
            {
                Assert.AreEqual((sides & Left) != 0 ? color : XlsxColor.Automatic, border.LeftBorderColor);
                Assert.AreEqual((sides & Right) != 0 ? color : XlsxColor.Automatic, border.RightBorderColor);
                Assert.AreEqual((sides & Top) != 0 ? color : XlsxColor.Automatic, border.TopBorderColor);
                Assert.AreEqual((sides & Bottom) != 0 ? color : XlsxColor.Automatic, border.BottomBorderColor);
            }
        }

        private static IEnumerable<object> BorderApiSetters()
        {
            var styleValues = EnumPolyfill.GetValues<XlsxBorderStyleValues>();
            var colors = new[] { XlsxColor.Red, XlsxColor.Black, XlsxColor.Automatic };

            // Outside border style - check that once set, all outer borders are set to the style
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.OutsideBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorder, (border, value) => border.OutsideBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorder, (border, value) => border.OutsideBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.OutsideBorder = value, styleValues);

            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.SetOutsideBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorder, (border, value) => border.SetOutsideBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorder, (border, value) => border.SetOutsideBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.SetOutsideBorder(value), styleValues);

            // Outside border color - check that once set, all outer borders are set to the color
            // Because of hash key and repos, we must first set style other than none
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);

            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetOutsideBorder(XlsxBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);

            // Left border
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.LeftBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.SetLeftBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetLeftBorder(XlsxBorderStyleValues.Thin).Border.LeftBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetLeftBorder(XlsxBorderStyleValues.Thin).Border.SetLeftBorderColor(value), colors);

            // Right border
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorder, (border, value) => border.RightBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorder, (border, value) => border.SetRightBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetRightBorder(XlsxBorderStyleValues.Thin).Border.RightBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetRightBorder(XlsxBorderStyleValues.Thin).Border.SetRightBorderColor(value), colors);

            // Top border
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorder, (border, value) => border.TopBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorder, (border, value) => border.SetTopBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetTopBorder(XlsxBorderStyleValues.Thin).Border.TopBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetTopBorder(XlsxBorderStyleValues.Thin).Border.SetTopBorderColor(value), colors);

            // Bottom border
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.BottomBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.SetBottomBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetBottomBorder(XlsxBorderStyleValues.Thin).Border.BottomBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetBottomBorder(XlsxBorderStyleValues.Thin).Border.SetBottomBorderColor(value), colors);

            // Diagonal up
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalUp, (border, value) => border.DiagonalUp = value, true, false);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalUp, (border, value) => border.SetDiagonalUp(value), true, false);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalUp, (border, _) => border.SetDiagonalUp(), true);

            // Diagonal down
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalDown, (border, value) => border.DiagonalDown = value, true, false);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalDown, (border, value) => border.SetDiagonalDown(value), true, false);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalDown, (border, _) => border.SetDiagonalDown(), true);

            // Diagonal border
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalBorder, (border, value) => border.DiagonalBorder = value, styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalBorder, (border, value) => border.SetDiagonalBorder(value), styleValues);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalBorderColor, (border, value) => border.SetDiagonalBorder(XlsxBorderStyleValues.Thin).Border.DiagonalBorderColor = value, colors);
            yield return FormatTestCase<IXlsxBorder>.ForBorder(border => border.DiagonalBorderColor, (border, value) => border.SetDiagonalBorder(XlsxBorderStyleValues.Thin).Border.SetDiagonalBorderColor(value), colors);
        }

        private static IEnumerable<object> BorderColorSetters()
        {
            yield return MakeTestCase(border => border.LeftBorderColor, (border, value) => border.LeftBorderColor = value, (border, style) => border.LeftBorder = style);
            yield return MakeTestCase(border => border.TopBorderColor, (border, value) => border.TopBorderColor = value, (border, style) => border.TopBorder = style);
            yield return MakeTestCase(border => border.RightBorderColor, (border, value) => border.RightBorderColor = value, (border, style) => border.RightBorder = style);
            yield return MakeTestCase(border => border.BottomBorderColor, (border, value) => border.BottomBorderColor = value, (border, style) => border.BottomBorder = style);
            yield return MakeTestCase(border => border.DiagonalBorderColor, (border, value) => border.DiagonalBorderColor = value, (border, style) => border.DiagonalBorder = style);
            yield break;

            static TestCaseData MakeTestCase(Func<IXlsxBorder, XlsxColor> getColor, Action<IXlsxBorder, XlsxColor> setColor, Action<IXlsxBorder, XlsxBorderStyleValues> setStyle)
            {
                return new TestCaseData(getColor, setColor, setStyle);
            }
        }
    }
}
