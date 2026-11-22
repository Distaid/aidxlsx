using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.PivotTables;

[TestFixture]
public class XlsxPivotTableFiltersTests
{
    [Test]
    public void Adding_and_removing_filters_shifts_pivot_table_area()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var data = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "City", "Flavor", "Value"),
            ("Cake", "Tokyo", "Vanilla", 7),
        });

        var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), data);

        // No filter, the table is at the original cell
        Assert.AreEqual("E2", ((XlsxPivotTable)pt).Area.ToString());

        pt.ReportFilters.Add("City");

        // First filter also adds divider row between filter and the table.
        Assert.AreEqual("E4", ((XlsxPivotTable)pt).Area.ToString());

        pt.ReportFilters.Add("Flavor");

        // When second filter is added, there is no need to add second divider row.
        Assert.AreEqual("E5", ((XlsxPivotTable)pt).Area.ToString());

        pt.ReportFilters.Remove("City");
        Assert.AreEqual("E4", ((XlsxPivotTable)pt).Area.ToString());

        pt.ReportFilters.Remove("Flavor");
        Assert.AreEqual("E2", ((XlsxPivotTable)pt).Area.ToString());
    }
}
