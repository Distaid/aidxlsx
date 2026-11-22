using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.PivotTables.PivotStyleFormats;
using Aid.Xlsx.Model.PivotTables.PivotValues;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.PivotTables.Style;

[TestFixture]
internal class XlsxPivotTableStyleFormatsTests
{
    [Test]
    public void Add_grand_row_total_styles()
    {
        TestHelper.CreateAndCompare(wb =>
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Price"),
                ("Cake", 9),
                ("Pie", 7),
                ("Cake", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            ptSheet.Column("A").Width = 15;
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.RowLabels.Add("Name");
            pt.Values.Add("Price", "Avg $").SetSummaryFormula(XlsxPivotSummary.Average);
            pt.Values.Add("Price", "Max $").SetSummaryFormula(XlsxPivotSummary.Maximum);

            pt.StyleFormats.RowGrandTotalFormats
                .ForElement(XlsxPivotStyleFormatElement.All).Style
                .Font.SetFontSize(15)
                .Font.SetUnderline(XlsxFontUnderlineValues.Double);
            pt.StyleFormats.RowGrandTotalFormats
                .ForElement(XlsxPivotStyleFormatElement.Label).Style
                .Font.SetFontColor(XlsxColor.Green);
            pt.StyleFormats.RowGrandTotalFormats
                .ForElement(XlsxPivotStyleFormatElement.Data).Style
                .Font.SetFontColor(XlsxColor.Red);
        }, @"Other\PivotTable\Style\Add_grand_row_total_styles.xlsx");
    }

    [Test]
    public void Add_grand_column_total_styles()
    {
        TestHelper.CreateAndCompare(wb =>
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Month", "Price"),
                ("Cake", "Jan", 9),
                ("Pie", "Jan", 7),
                ("Cake", "Feb", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            ptSheet.Column("A").Width = 15;
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.RowLabels.Add("Name");
            pt.RowLabels.Add("Month");
            pt.Values.Add("Price");

            pt
                .SetShowGrandTotalsColumns(true)
                .SetShowGrandTotalsRows(false);

            pt.StyleFormats.ColumnGrandTotalFormats
                .ForElement(XlsxPivotStyleFormatElement.All).Style
                .Font.SetFontSize(15)
                .Font.SetUnderline(XlsxFontUnderlineValues.Double);
            pt.StyleFormats.ColumnGrandTotalFormats
                .ForElement(XlsxPivotStyleFormatElement.Label).Style
                .Font.SetFontColor(XlsxColor.Green);
            pt.StyleFormats.ColumnGrandTotalFormats
                .ForElement(XlsxPivotStyleFormatElement.Data).Style
                .Font.SetFontColor(XlsxColor.Red);
        }, @"Other\PivotTable\Style\Add_grand_column_total_styles.xlsx");
    }
}

