using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.PivotTables;

/// <summary>
/// Tests methods of interface <see cref="IXlsxPivotField"/> implemented through <see cref="XlsxPivotTableAxisField"/>.
/// </summary>
[TestFixture]
internal class XlsxPivotTableAxisFieldTests
{
    [Test]
    public void CustomName_can_be_changed()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), range);
        var colorField = pt.RowLabels.Add("Color");

        colorField.SetCustomName("Changed color");

        Assert.AreEqual("Changed color", pt.RowLabels.Get(0).CustomName);
    }

    [Test]
    public void CustomName_throws_exception_when_name_is_already_used()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("ID", "Color", "Count"),
            (1, "Blue", 10),
        });
        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), range);
        var idField = pt.RowLabels.Add("ID", "Custom ID");
        var colorField = pt.RowLabels.Add("Color");

        var ex1 = Assert.Throws<ArgumentException>(() => idField.SetCustomName("Color"))!;
        Assert.AreEqual("Custom name 'Color' is already used by another field.", ex1.Message);
        var ex2 = Assert.Throws<ArgumentException>(() => colorField.SetCustomName("Custom ID"));
        Assert.AreEqual("Custom name 'Custom ID' is already used by another field.", ex2.Message);
    }
}
