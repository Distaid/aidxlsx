using System.Data;
using Aid.Xlsx;

using var file = new MemoryStream();
using (var workbook = new XLWorkbook())
{
    var sheet = workbook.Worksheets.Add("Sales");
    using var sales = new DataTable();
    sales.Columns.Add("Item", typeof(string));
    sales.Columns.Add("Amount", typeof(double));
    sales.Rows.Add("Apples", 20d);
    sales.Rows.Add("Pears", 22d);
    sheet.Cell("A1").InsertTable(sales, "SalesData");
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
Require(salesSheet.Cell("D1").GetDouble() == 42d, "evaluated formula");
Require(salesSheet.Cell("D1").Style.NumberFormat.Format == "0.00", "number format");
Require(salesSheet.Cell("D2").GetDateTime() == new DateTime(2026, 1, 2), "date value");
Require(salesSheet.Cell("D3").GetComment().Text == "Verified", "comment");
Console.WriteLine("Workbook round-trip passed.");

static void Require(bool valid, string scenario)
{
    if (!valid)
        throw new InvalidOperationException($"Workbook round-trip failed for {scenario}.");
}
