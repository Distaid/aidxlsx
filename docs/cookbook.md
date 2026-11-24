# Cookbook

Complete recipes for common workbook tasks. Each recipe has a goal, a runnable
top-level program, and notes that link to the detailed guide.

## Monthly report with styles and totals

**Goal:** build a report with a styled header, money formats, totals, and sized
columns.

```csharp
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("January");

sheet.Cell("A3").Value = "Region";
sheet.Cell("B3").Value = "Units";
sheet.Cell("C3").Value = "Revenue";
sheet.Cell("A4").Value = "North";
sheet.Cell("B4").Value = 120;
sheet.Cell("C4").Value = 4800m;
sheet.Cell("A5").Value = "South";
sheet.Cell("B5").Value = 95;
sheet.Cell("C5").Value = 3800m;

var header = sheet.Range("A3:C3");
header.Style.Font.Bold = true;
header.Style.Font.FontColor = XlsxColor.White;
header.Style.Fill.BackgroundColor = XlsxColor.DarkBlue;
sheet.Range("C4:C5").Style.NumberFormat.Format = "#,##0.00";

sheet.Range("A6:C6").Style.Font.Bold = true;
sheet.Cell("A6").Value = "Total";
sheet.Cell("B6").FormulaA1 = "SUM(B4:B5)";
sheet.Cell("C6").FormulaA1 = "SUM(C4:C5)";
sheet.Range("C6").Style.NumberFormat.Format = "#,##0.00";
sheet.Columns(1, 3).AdjustToContents();

workbook.SaveAs("january-sales.xlsx", new SaveOptions
{
    EvaluateFormulasBeforeSaving = true
});
```

- Cached formula values are omitted by default; the save option stores results.
- `AdjustToContents` depends on font metrics, which can differ on Linux.

Guide: [Formatting and worksheet features](formatting-and-worksheets.md).

## Import CSV or DataTable and create a table

**Goal:** load delimited text into a `DataTable` and write it as a styled Excel table.

```csharp
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

var lines = File.ReadAllLines("sales.csv");
var data = new DataTable("Sales");
data.Columns.Add("Region", typeof(string));
data.Columns.Add("Units", typeof(int));

foreach (var line in lines.Skip(1))
{
    var parts = line.Split(',');
    data.Rows.Add(parts[0], int.Parse(parts[1], CultureInfo.InvariantCulture));
}

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");
var table = sheet.Cell("A1").InsertTable(data, "SalesTable");
table.Theme = XlsxTableTheme.TableStyleMedium2;

workbook.SaveAs("sales-table.xlsx");
```

- The split handles simple CSV without quotes or line breaks; use a CSV parser for
  general input, then fill the `DataTable`.

Guide: [Tables and data import](tables-and-data-import.md).

## Highlight values with conditional formatting

**Goal:** color cells by threshold without changing their values.

```csharp
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");
sheet.Cell("A1").Value = "Item";
sheet.Cell("B1").Value = "Revenue";
sheet.Cell("A2").Value = "Pears";
sheet.Cell("B2").Value = 6200m;
sheet.Cell("A3").Value = "Apples";
sheet.Cell("B3").Value = 2800m;

var revenue = sheet.Range("B2:B3");
revenue.AddConditionalFormat()
    .WhenGreaterThan(5000)
    .Fill.SetBackgroundColor(XlsxColor.LightGreen);
revenue.AddConditionalFormat()
    .WhenLessThan(3000)
    .Fill.SetBackgroundColor(XlsxColor.LightYellow);

workbook.SaveAs("conditional-format.xlsx");
```

- Excel or LibreOffice evaluates the rules; Aid.Xlsx only stores them.

Guide: [Formatting and worksheet features](formatting-and-worksheets.md).

## Add a data-validation dropdown

**Goal:** restrict a column to a list and show input and error messages.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Orders");
sheet.Cell("A1").Value = "Order";
sheet.Cell("B1").Value = "Status";

var status = sheet.Range("B2:B200").CreateDataValidation();
status.List("Pending,Paid,Cancelled");
status.InputTitle = "Status";
status.InputMessage = "Choose Pending, Paid, or Cancelled.";
status.ErrorTitle = "Invalid status";
status.ErrorMessage = "Pick one of the three allowed values.";
status.ShowInputMessage = true;
status.ShowErrorMessage = true;

workbook.SaveAs("orders-validation.xlsx");
```

- `List(string)` splits on commas; `List(IXlsxRange)` reads values from cells.
- Validation limits interactive editing only. Validate values written by code too.

Guide: [Formatting and worksheet features](formatting-and-worksheets.md).

## Freeze the header and set up printing

**Goal:** keep the header visible and print a landscape page with repeated headings.

```csharp
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Register");
sheet.Cell("A1").Value = "Date";
sheet.Cell("B1").Value = "Item";
sheet.Cell("C1").Value = "Amount";
sheet.Cell("B2").Value = "Pears";
sheet.Cell("C2").Value = 12.50m;

sheet.SheetView.FreezeRows(1);
sheet.PageSetup.PrintAreas.Add("A1:D40");
sheet.PageSetup.SetRowsToRepeatAtTop(1, 1);
sheet.PageSetup.PageOrientation = XlsxPageOrientation.Landscape;
sheet.PageSetup.PaperSize = XlsxPaperSize.A4Paper;
sheet.PageSetup.FitToPages(1, 1);

workbook.SaveAs("register-print.xlsx");
```

- `FreezeRows(1)`, `FreezeColumns(2)`, and `Freeze(rows, columns)` control panes.

Guide: [Formatting and worksheet features](formatting-and-worksheets.md).

## Place an image at a cell anchor

**Goal:** add a logo and attach it to a cell so it moves with the content.

```csharp
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");
sheet.Cell("A1").Value = "Quarterly report";

var picture = sheet.AddPicture("logo.png", "Logo");
picture.MoveTo(sheet.Cell("A3"));
picture.WithSize(160, 48);
picture.WithPlacement(XlsxPicturePlacement.MoveAndSize);

workbook.SaveAs("report-with-logo.xlsx");
```

- `AddPicture(path, name)` infers the format; the stream overload accepts bytes.
- `MoveTo(cell)` anchors the top-left corner; offset overloads position inside a cell.

Guide: [Formatting and worksheet features](formatting-and-worksheets.md).

## Use defined names in formulas

**Goal:** name a data range and a rate, then use both in formulas.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Invoice");
sheet.Cell("A2").Value = "Pears";
sheet.Cell("B2").Value = 120.50m;
sheet.Cell("A3").Value = "Apples";
sheet.Cell("B3").Value = 80.00m;
sheet.Cell("F1").Value = 0.20m;

workbook.DefinedNames.Add("LineItems", sheet.Range("B2:B3"));
sheet.DefinedNames.Add("TaxRate", sheet.Range("F1"));

sheet.Cell("B4").FormulaA1 = "SUM(LineItems)";
sheet.Cell("B5").FormulaA1 = "B4*(1+TaxRate)";

workbook.SaveAs("invoice-names.xlsx", new SaveOptions
{
    EvaluateFormulasBeforeSaving = true
});
```

- Workbook names resolve from every sheet; sheet names resolve on their own sheet
  first and need `'Invoice'!TaxRate` elsewhere.

Guide: [Workbook files and streams](workbook-files.md).

## Convert a legacy .xls workbook

**Goal:** convert an Excel 97-2003 file and report unsupported input cleanly.

```csharp
using System;
using System.IO;
using Aid.Xlsx.Conversion;

Directory.CreateDirectory("converted");

try
{
    XlsConverter.Convert("legacy.xls");
    XlsConverter.Convert("legacy.xls", "converted/legacy.xlsx");
}
catch (NotSupportedException ex)
{
    Console.WriteLine($"Unsupported workbook: {ex.Message}");
}
catch (InvalidDataException ex)
{
    Console.WriteLine($"Malformed workbook: {ex.Message}");
}
```

- The one-argument overload writes `legacy.xlsx` next to the source. The two-argument
  overload needs a different path and an existing directory.
- The converter carries values, formatting, dimensions, merges, and its supported
  formula subset; other formulas become static values. Encrypted, pre-BIFF8, chart,
  macro, and dialog-sheet input is rejected.

Guide: [Workbook files and streams](workbook-files.md) and
[Workbook compatibility](compatibility.md).

## Stream a workbook from ASP.NET

**Goal:** generate a workbook in memory and return it as a download.

```csharp
using System.IO;
using Aid.Xlsx.Model.Workbooks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.MapGet("/reports/sales.xlsx", () =>
{
    using var workbook = new XlsxWorkbook();
    var sheet = workbook.Worksheets.Add("Sales");
    sheet.Cell("A1").Value = "Item";
    sheet.Cell("B1").Value = "Units";
    sheet.Cell("A2").Value = "Pears";
    sheet.Cell("B2").Value = 12;

    using var output = new MemoryStream();
    workbook.SaveAs(output);
    return Results.File(
        output.ToArray(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "sales.xlsx");
});

app.Run();
```

- A loaded workbook's `SaveAs(Stream)` needs a readable, seekable, writable stream.
  A response body usually is not; `MemoryStream` and `ToArray()` are portable.
- The workbook is held in memory; use a temporary file for very large reports.

Guide: [Workbook files and streams](workbook-files.md).

## NativeAOT-safe import

**Goal:** insert typed rows through a generated mapper with no reflection path.

```csharp
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Stock");
var rows = new[] { new StockItem("Pears", 12), new StockItem("Apples", 18) };

sheet.Cell("A1").InsertMappedTable(rows);
workbook.SaveAs("stock.xlsx");

[XlsxTable("StockData")]
internal sealed record StockItem(
    [property: XlsxColumn(Header = "Item", Order = 1)] string Name,
    [property: XlsxColumn(Order = 2)] int Units);
```

- `InsertMappedTable` creates a table; `InsertMappedData(rows, addHeadings: true)`
  writes a plain range.
- `InsertTable<T>` and `InsertData(IEnumerable)` use reflection and warn when
  trimmed. Do not use them on an AOT path.

Guide: [NativeAOT and trimming](native-aot.md).

## Protect a worksheet

**Goal:** leave an input range editable while protecting formulas and layout.

```csharp
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Input");
sheet.Cell("A1").Value = "Name";
sheet.Cell("B1").Value = "Amount";
sheet.Cell("C1").Value = "Total";
sheet.Cell("A2").Value = "Pears";
sheet.Cell("B2").Value = 12.50m;
sheet.Cell("C3").FormulaA1 = "B2*2";

sheet.Range("B2:B20").Style.Protection.Locked = false;
sheet.Protect(
    "change-me",
    XlsxProtectionAlgorithm.Algorithm.SHA512,
    XlsxSheetProtectionElements.SelectLockedCells |
    XlsxSheetProtectionElements.SelectUnlockedCells);

workbook.SaveAs("protected-input.xlsx");
```

- Cells are locked by default; unlock the input range before calling `Protect`.

Guide: [Formatting and worksheet features](formatting-and-worksheets.md).

## Add comments and rich text

**Goal:** attach a visible note and format part of a cell's text with multiple runs.

```csharp
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Notes");
sheet.Cell("B2").Value = 1250.75m;
sheet.Cell("B2").Style.NumberFormat.Format = "#,##0.00";

var comment = sheet.Cell("B2").GetComment();
comment.SetAuthor("Finance");
comment.SetVisible();
comment.AddText("Checked against the ledger.");

var rich = sheet.Cell("A1").CreateRichText();
rich.AddText("Q3 ").SetBold();
rich.AddText("revenue").SetFontColor(XlsxColor.Red);

workbook.SaveAs("annotations.xlsx");
```

- `GetComment()` creates a note when none exists; check `HasComment` to read only.
- `SetVisible()` shows the note on open; the parameterless overload is correct.
- `CreateRichText()` starts from existing cell text; add runs in reading order.

Guide: [Formatting and worksheet features](formatting-and-worksheets.md).
