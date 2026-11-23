using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.SpreadsheetML.Workbooks;

[TestFixture]
public class TemplateExportTests
{
    [Test]
    public void TemplateEditsPreserveMergesCopiedStyleAndRowHeights()
    {
        using var template = new MemoryStream();
        var fill = XlsxColor.FromHtml("#336699");
        var border = XlsxColor.FromHtml("#AABBCC");
        const string numberFormat = "0.0000";

        using (var workbook = new XlsxWorkbook())
        {
            var sheet = workbook.AddWorksheet("Export");
            sheet.Range("C1:D1").Merge().FirstCell().Value = "Header";
            sheet.Range("B5:D5").Merge().FirstCell().Value = "Shifted total";
            sheet.Range("E5:F5").Merge().FirstCell().Value = "Unchanged note";
            sheet.Cell("A8").Value = "Footer";
            sheet.Row(8).Height = 41;

            var source = sheet.Cell("B2");
            source.Style.Fill.BackgroundColor = fill;
            source.Style.Border.LeftBorder = XlsxBorderStyleValues.Thick;
            source.Style.Border.LeftBorderColor = border;
            source.Style.NumberFormat.Format = numberFormat;
            workbook.SaveAs(template, validate: true);
        }

        template.Position = 0;
        using var output = new MemoryStream();
        using (var workbook = new XlsxWorkbook(template))
        {
            var sheet = workbook.Worksheet("Export");
            sheet.Range("B4:D4").InsertRowsBelow(2);
            sheet.Range("C4:C6").Style = sheet.Cell("B2").Style;
            sheet.Cell("C5").Value = 12.5;
            sheet.Range("7:7").InsertRowsBelow(1);
            sheet.Row(8).Height = 27;
            workbook.SaveAs(output, validate: true);
        }

        output.Position = 0;
        using var saved = new XlsxWorkbook(output);
        var export = saved.Worksheet("Export");
        var styledCell = export.Cell("C5");
        Assert.Multiple(() =>
        {
            Assert.That(export.MergedRanges.Select(range => range.RangeAddress.ToString()),
                Is.EquivalentTo(new[] { "C1:D1", "B7:D7", "E5:F5" }));
            Assert.That(export.Cell("C1").GetText(), Is.EqualTo("Header"));
            Assert.That(export.Cell("B7").GetText(), Is.EqualTo("Shifted total"));
            Assert.That(export.Cell("E5").GetText(), Is.EqualTo("Unchanged note"));
            Assert.That(styledCell.GetDouble(), Is.EqualTo(12.5));
            Assert.That(styledCell.Style.Fill.BackgroundColor, Is.EqualTo(fill));
            Assert.That(styledCell.Style.Border.LeftBorder, Is.EqualTo(XlsxBorderStyleValues.Thick));
            Assert.That(styledCell.Style.Border.LeftBorderColor, Is.EqualTo(border));
            Assert.That(styledCell.Style.NumberFormat.Format, Is.EqualTo(numberFormat));
            Assert.That(export.Row(8).Height, Is.EqualTo(27));
            Assert.That(export.Row(9).Height, Is.EqualTo(41));
            Assert.That(export.Cell("A9").GetText(), Is.EqualTo("Footer"));
        });
    }

    [Test]
    public void WorksheetCellsUsedSupportsBothPredicateOverloads()
    {
        using var workbook = new XlsxWorkbook();
        var sheet = workbook.AddWorksheet("Export");
        sheet.Cell("A1").Value = "Invoice";
        sheet.Cell("A2").Value = 42;
        sheet.Cell("A3").Style.Font.Bold = true;

        var texts = sheet.CellsUsed(cell => cell.DataType == XlsxDataType.Text);
        var boldCells = sheet.CellsUsed(XlsxCellsUsedOptions.AllContents | XlsxCellsUsedOptions.NormalFormats,
            cell => cell.Style.Font.Bold);

        Assert.Multiple(() =>
        {
            Assert.That(texts.Select(cell => cell.GetText()), Is.EqualTo(new[] { "Invoice" }));
            Assert.That(boldCells.Select(cell => cell.Address.ToString()), Is.EqualTo(new[] { "A3" }));
        });
    }

    [Test]
    public void NewWorksheetsInheritAndRetainTheWorkbookDefaultStyle()
    {
        using var output = new MemoryStream();
        var fill = XlsxColor.FromHtml("#336699");
        const string numberFormat = "0.0000";

        using (var workbook = new XlsxWorkbook())
        {
            workbook.Style.Font.FontName = "Arial";
            workbook.Style.Alignment.Horizontal = XlsxAlignmentHorizontalValues.Center;
            workbook.Style.Fill.BackgroundColor = fill;
            workbook.Style.NumberFormat.Format = numberFormat;

            var sheet = workbook.AddWorksheet("Export");
            sheet.Cell("A1").Value = 12.5;
            sheet.Cell("B1").Value = 13.5;
            sheet.Cell("B1").Style.Font.FontName = "Tahoma";

            AssertDefaultStyle(workbook.Style);
            AssertDefaultStyle(sheet.Style);
            AssertDefaultStyle(sheet.Cell("A1").Style);
            workbook.SaveAs(output, validate: true);
        }

        output.Position = 0;
        using var saved = new XlsxWorkbook(output);
        var export = saved.Worksheet("Export");
        AssertDefaultStyle(saved.Style);
        AssertDefaultStyle(export.Cell("A1").Style);
        Assert.That(export.Cell("B1").Style.Font.FontName, Is.EqualTo("Tahoma"));

        void AssertDefaultStyle(IXlsxStyle style)
        {
            Assert.Multiple(() =>
            {
                Assert.That(style.Font.FontName, Is.EqualTo("Arial"));
                Assert.That(style.Alignment.Horizontal, Is.EqualTo(XlsxAlignmentHorizontalValues.Center));
                Assert.That(style.Fill.BackgroundColor, Is.EqualTo(fill));
                Assert.That(style.NumberFormat.Format, Is.EqualTo(numberFormat));
            });
        }
    }
}
