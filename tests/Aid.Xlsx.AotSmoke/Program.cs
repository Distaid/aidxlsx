using Aid.Xlsx.Conversion;
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

using var file = new MemoryStream();

using (var workbook = new XlsxWorkbook())
{
    workbook.Use1904DateSystem = true;
    var sheet = workbook.Worksheets.Add("Sales");
    var sales = new[] { new Sale("Apples", 20m, new DateOnly(2026, 1, 2)), new Sale("Pears", 22m, null) };
    sheet.Cell("A1").InsertMappedTable(sales).Theme = XlsxTableTheme.TableStyleDark2;
    sheet.Cell("F1").InsertMappedData(sales, addHeadings: true);
    sheet.Cell("D1").FormulaA1 = "SUM(SalesData[Amount])";
    sheet.Cell("D1").Style.NumberFormat.Format = "0.00";
    sheet.Cell("D2").Value = new DateTime(2026, 1, 2);
    sheet.Cell("D2").Style.DateFormat.Format = "yyyy-mm-dd";
    sheet.Cell("D3").GetComment().AddText("Verified");
    sheet.Range("D5:E5").Merge().FirstCell().Value = "Merged";
    sheet.Range("F2:F3").CreateDataValidation().TextLength.Between(1, 20);
    sheet.Cell("D6").Value = "Link";
    sheet.Cell("D6").SetHyperlink(new XlsxHyperlink(new Uri("https://example.org/sales")));
    sheet.Range("B2:B3").AddConditionalFormat().WhenGreaterThan(20).Fill.SetBackgroundColor(XlsxColor.Green);
    sheet.Protect("sheet password", XlsxProtectionAlgorithm.Algorithm.SHA512);
    workbook.Protect("workbook password", XlsxProtectionAlgorithm.Algorithm.SHA512);
    workbook.SaveAs(file, new SaveOptions { EvaluateFormulasBeforeSaving = true, ValidatePackage = true });
}

file.Position = 0;
using var reloaded = new XlsxWorkbook(file);
var salesSheet = reloaded.Worksheet("Sales");
Require(salesSheet.Table("SalesData").Field("Amount").Name == "Amount", "table heading");
Require(ReferenceEquals(salesSheet.Table("SalesData").Theme, XlsxTableTheme.TableStyleDark2), "generated table theme lookup");
Require(salesSheet.Cell("A2").GetString() == "Apples", "shared string");
Require(salesSheet.Cell("B2").Style.NumberFormat.Format == "0.000", "generated column format");
Require(salesSheet.Cell("C2").GetDateTime() == new DateTime(2026, 1, 2), "generated DateOnly value");
Require(salesSheet.Cell("C3").IsEmpty(), "generated nullable value");
Require(salesSheet.Cell("F1").GetString() == "Item" && salesSheet.Cell("G3").GetDouble() == 22d, "generated data range");
Require(salesSheet.Cell("D1").GetDouble() == 42d, "evaluated formula");
Require(salesSheet.Cell("D1").Style.NumberFormat.Format == "0.00", "number format");
Require(reloaded.Use1904DateSystem && salesSheet.Cell("D2").GetDateTime() == new DateTime(2026, 1, 2), "1904 date system");
Require(salesSheet.Cell("D3").GetComment().Text == "Verified", "comment");
Require(salesSheet.Cell("E5").IsMerged() && salesSheet.Cell("D5").GetString() == "Merged", "merged range");
Require(salesSheet.Cell("D6").GetHyperlink().ExternalAddress == new Uri("https://example.org/sales"), "hyperlink");
var validation = salesSheet.Cell("F2").GetDataValidation();
Require(validation.AllowedValues == XlsxAllowedValues.TextLength && validation.MinValue == "1" && validation.MaxValue == "20", "data validation");
var conditionalFormat = salesSheet.ConditionalFormats.Single();
Require(conditionalFormat.Operator == XlsxCFOperator.GreaterThan && conditionalFormat.Style.Fill.BackgroundColor == XlsxColor.Green, "conditional formatting");
Require(salesSheet.IsPasswordProtected && reloaded.IsPasswordProtected, "SHA-512 protection");
salesSheet.Unprotect("sheet password");
reloaded.Unprotect("workbook password");

using var legacy = typeof(Sale).Assembly.GetManifestResourceStream("legacy.xls")
    ?? throw new InvalidOperationException("The BIFF8 fixture is missing.");
using var converted = new MemoryStream();
XlsConverter.Convert(legacy, converted);
converted.Position = 0;
using var convertedWorkbook = new XlsxWorkbook(converted);
Require(convertedWorkbook.Worksheet("Hidden").Cell("A1").GetDouble() == 42, "BIFF8 conversion");
using var formulaSource = typeof(Sale).Assembly.GetManifestResourceStream("formulas.xls")
    ?? throw new InvalidOperationException("The BIFF8 formula fixture is missing.");
using var formulaOutput = new MemoryStream();
XlsConverter.Convert(formulaSource, formulaOutput);
formulaOutput.Position = 0;
using var formulaWorkbook = new XlsxWorkbook(formulaOutput);
var formulaCell = formulaWorkbook.Worksheet("Data").Cell("B1");
Require(formulaCell.FormulaA1 == "SUM(A1:$A$2)*2+1" && formulaCell.CachedValue.GetNumber() == 21, "BIFF8 formula conversion");
if (args.Length == 1)
{
    using var artifact = File.Create(args[0]);
    file.WriteTo(artifact);
}
Console.WriteLine("Workbook round-trip passed.");

static void Require(bool valid, string scenario)
{
    if (!valid)
        throw new InvalidOperationException($"Workbook round-trip failed for {scenario}.");
}

[XlsxTable("SalesData")]
internal sealed record Sale(
    [property: XlsxColumn(Header = "Item", Order = 1)] string Product,
    [property: XlsxColumn(Order = 2, Format = "0.000")] decimal Amount,
    DateOnly? SoldOn)
{
    [XlsxIgnore]
    public Uri? Link { get; init; }
}
