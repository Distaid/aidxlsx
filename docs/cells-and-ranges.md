# Cells and ranges

Use A1 notation or numeric row and column indexes to address a worksheet cell. In
worksheet coordinates, row `1` and column `1` identify `A1`. Address a rectangular
area with `Range("A1:C3")`.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Inventory");

sheet.Cell("A1").Value = "SKU";
sheet.Cell("B1").Value = "Units";
sheet.Cell("C1").Value = "Received";
sheet.Cell(2, 1).Value = "00123";
sheet.Cell(2, 2).Value = 12;
sheet.Cell(2, 3).Value = new DateTime(2026, 1, 15);

string sku = sheet.Cell("A2").GetText();
DateTime received = sheet.Cell("C2").GetDateTime();

if (sheet.Cell("B2").TryGetValue<int>(out var units))
    Console.WriteLine($"{sku}: {units} units received on {received:yyyy-MM-dd}");

foreach (var cell in sheet.Range("A1:C10").CellsUsed())
    Console.WriteLine($"{cell.Address}: {cell.GetFormattedString()}");

workbook.SaveAs("inventory.xlsx");
```

- Assign identifiers with leading zeros as strings. Assign quantities and dates as
  numbers and `DateTime` values to use them in formulas and filters.
- Use `GetText()`, `GetDouble()`, and `GetDateTime()` when you know the cell type.
  Use `TryGetValue<T>()` when a conversion may fail. Inspect `DataType` when reading
  mixed data.
- Use `GetFormattedString()` for display text based on the cell's number format.
  Pass a `CultureInfo` when your application needs a specific display culture.
- Use `CellsUsed()` to visit populated cells in a range or worksheet. Use
  `Cells()` on a bounded range when you also need to visit empty positions.

## Merge cells

Call `Merge()` on a range and put the value in its upper-left cell. Use an empty
range when adding a title so that merging does not discard existing values.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");

sheet.Range("A1:C1").Merge();
sheet.Cell("A1").Value = "January inventory";

workbook.SaveAs("report.xlsx");
```
