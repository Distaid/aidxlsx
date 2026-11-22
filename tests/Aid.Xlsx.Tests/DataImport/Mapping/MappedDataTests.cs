using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.DataImport.Mapping;

public class MappedDataTests
{
    [Test]
    public void InsertsGeneratedColumnsInExplicitOrderAndSkipsIgnoredMembers()
    {
        using var workbook = new XlsxWorkbook();
        var sheet = workbook.AddWorksheet("Orders");
        var table = sheet.Cell("B3").InsertMappedTable([
            new OrderRow("Apple", 12.5m, "Internal"),
            new OrderRow("Pear", null, "Internal")]);

        Assert.That(table.Name, Is.EqualTo("OrdersData"));
        Assert.That(table.Fields.Select(field => field.Name), Is.EqualTo(new[] { "Price", "Product name" }));
        Assert.That(sheet.Cell("B4").GetDouble(), Is.EqualTo(12.5));
        Assert.That(sheet.Cell("B4").Style.NumberFormat.Format, Is.EqualTo("0.000"));
        Assert.That(sheet.Cell("B5").Value.IsBlank, Is.True);
        Assert.That(sheet.Cell("C5").GetString(), Is.EqualTo("Pear"));
        Assert.That(sheet.Cell("D4").Value.IsBlank, Is.True);
    }

    [Test]
    public void MissingMapperFailsBeforeWritingCells()
    {
        using var workbook = new XlsxWorkbook();
        var sheet = workbook.AddWorksheet("Orders");
        var cell = sheet.Cell("A1");
        cell.Value = "Existing";

        Assert.That(() => cell.InsertMappedData(new[] { new UnmappedRow("Apple") }, addHeadings: true),
            Throws.TypeOf<NotSupportedException>());
        Assert.That(cell.GetString(), Is.EqualTo("Existing"));
        Assert.That(sheet.Cell("A2").Value.IsBlank, Is.True);
    }

    private sealed record UnmappedRow(string Product);
}

[XlsxTable("OrdersData")]
internal sealed record OrderRow(
    [property: XlsxColumn(Header = "Product name", Order = 2)] string Product,
    [property: XlsxColumn(Order = 1, Format = "0.000")] decimal? Price,
    [property: XlsxIgnore] string InternalNote);
