# Tables and data import

## Create and filter a table from a range

Populate the header row and data cells, then call `CreateTable()` on the range.
With `AutoFilter.Column(1)`, the index is relative to the table's first column.
The example table begins in worksheet column B.

```csharp
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Catalog");

sheet.Cell("B1").Value = "Category";
sheet.Cell("C1").Value = "Item";
sheet.Cell("B2").Value = "Fruit";
sheet.Cell("C2").Value = "Pears";
sheet.Cell("B3").Value = "Vegetables";
sheet.Cell("C3").Value = "Carrots";

var table = sheet.Range("B1:C3").CreateTable("CatalogItems");
table.Theme = XlsxTableTheme.TableStyleMedium9;
table.ShowAutoFilter = true;
table.AutoFilter.Column(1).AddFilter("Fruit");

workbook.SaveAs("catalog.xlsx");
```

An auto-filter controls row visibility. The filtered-out rows remain in the workbook.
After opening the file, access the table by name with `workbook.Table("CatalogItems")`.
Use `DataRange` to work with its body and `Field("Item")` to select field metadata by name.

## Import a DataTable

`InsertTable(DataTable, string)` takes headings from the `DataTable` column names and
creates an Excel table starting at the selected cell. For a plain range, use
`InsertData(DataTable)` and write any desired headings separately.

```csharp
using System.Data;
using Aid.Xlsx.Model.Workbooks;

using var data = new DataTable();
data.Columns.Add("Item", typeof(string));
data.Columns.Add("Units", typeof(int));
data.Rows.Add("Pears", 12);
data.Rows.Add("Apples", 18);

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Stock");
sheet.Cell("A1").InsertTable(data, "StockItems");

workbook.SaveAs("stock.xlsx");
```

## Import typed records with a generated mapper

Annotate the row type with `[XlsxTable]`. The package's source generator maps public
instance fields and readable properties at build time. Use `XlsxColumn` for a
custom header, order, or number format and `XlsxIgnore` to omit a member.

```csharp
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");
var sales = new[] { new Sale("Pears", 12.50m), new Sale("Apples", 8m) };

sheet.Cell("A1").InsertMappedTable(sales);
workbook.SaveAs("mapped-sales.xlsx");

[XlsxTable("SalesData")]
internal record Sale(
    [property: XlsxColumn(Order = 0)] string Item,
    [property: XlsxColumn(Order = 1, Header = "Total", Format = "0.00")] decimal Amount);
```

For a plain range from the same record type, use `InsertMappedData(sales, addHeadings: true)`.
Use these generated-mapper methods on trim-sensitive paths. [NativeAOT and trimming](native-aot.md)
