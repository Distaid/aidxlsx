using Aid.Xlsx;
using Aid.Xlsx.Attributes;

using var file = new MemoryStream();
using (var workbook = new XLWorkbook())
{
    var sheet = workbook.Worksheets.Add("Sales");
    var sales = new[] { new Sale("Apples", 20m, new DateOnly(2026, 1, 2)), new Sale("Pears", 22m, null) };
    sheet.Cell("A1").InsertMappedTable(sales);
    sheet.Cell("F1").InsertMappedData(sales, addHeadings: true);
    sheet.Cell("D1").FormulaA1 = "SUM(SalesData[Amount])";
    sheet.Cell("D1").Style.NumberFormat.Format = "0.00";
    sheet.Cell("D2").Value = new DateTime(2026, 1, 2);
    sheet.Cell("D2").Style.DateFormat.Format = "yyyy-mm-dd";
    sheet.Cell("D3").GetComment().AddText("Verified");
    workbook.SaveAs(file, new SaveOptions { EvaluateFormulasBeforeSaving = true });
}

file.Position = 0;
using var reloaded = new XLWorkbook(file);
var salesSheet = reloaded.Worksheet("Sales");
Require(salesSheet.Table("SalesData").Field("Amount").Name == "Amount", "table heading");
Require(salesSheet.Cell("A2").GetString() == "Apples", "shared string");
Require(salesSheet.Cell("B2").Style.NumberFormat.Format == "0.000", "generated column format");
Require(salesSheet.Cell("C2").GetDateTime() == new DateTime(2026, 1, 2), "generated DateOnly value");
Require(salesSheet.Cell("C3").IsEmpty(), "generated nullable value");
Require(salesSheet.Cell("F1").GetString() == "Item" && salesSheet.Cell("G3").GetDouble() == 22d, "generated data range");
Require(salesSheet.Cell("D1").GetDouble() == 42d, "evaluated formula");
Require(salesSheet.Cell("D1").Style.NumberFormat.Format == "0.00", "number format");
Require(salesSheet.Cell("D2").GetDateTime() == new DateTime(2026, 1, 2), "date value");
Require(salesSheet.Cell("D3").GetComment().Text == "Verified", "comment");

using var legacy = typeof(Sale).Assembly.GetManifestResourceStream("legacy.xls")
    ?? throw new InvalidOperationException("The BIFF8 fixture is missing.");
using var converted = new MemoryStream();
XlsConverter.Convert(legacy, converted);
converted.Position = 0;
using var convertedWorkbook = new XLWorkbook(converted);
Require(convertedWorkbook.Worksheet("Hidden").Cell("A1").GetDouble() == 42, "BIFF8 conversion");
Console.WriteLine("Workbook round-trip passed.");

static void Require(bool valid, string scenario)
{
    if (!valid)
        throw new InvalidOperationException($"Workbook round-trip failed for {scenario}.");
}

[XlsxTable("SalesData")]
internal sealed record Sale(
    [property: XLColumn(Header = "Item", Order = 1)] string Product,
    [property: XLColumn(Order = 2, Format = "0.000")] decimal Amount,
    DateOnly? SoldOn)
{
    [XlsxIgnore]
    public Uri? Link { get; init; }
}
