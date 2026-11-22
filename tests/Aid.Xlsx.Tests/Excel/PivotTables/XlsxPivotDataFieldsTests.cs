using NUnit.Framework;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.PivotTables.PivotValues;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel.PivotTables;

/// <summary>
/// Test methods of interface <see cref="IXlsxPivotValues"/> implemented through <see cref="XlsxPivotDataFields"/> class.
/// </summary>
internal class XlsxPivotDataFieldsTests
{
    #region IXlsxPivotValues methods

    #region Add

    [Test]
    public void Add_source_name_must_be_from_pivot_cache_field_names()
    {
        using var wb = new XlsxWorkbook();
        var data = wb.AddWorksheet();
        var range = data.Cell("A1").InsertData(new object[]
        {
            ("Name", "Price"),
            ("Cake", 10),
        });
        var ptSheet = wb.AddWorksheet();
        var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => pt.Values.Add("Wrong field name"));

        Assert.NotNull(ex);
        StringAssert.StartsWith("Field 'Wrong field name' is not in the fields of a pivot cache. Should be one of 'Name','Price'.", ex.Message);
    }

    #endregion

    #endregion
}
