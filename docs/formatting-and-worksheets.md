# Formatting and worksheet features

Apply a style to one cell or to a range. Format numeric and date columns with an
Excel format code so that the values can still be used in calculations. Run
`AdjustToContents()` after filling and styling the cells.

```csharp
using System;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Orders");

sheet.Cell("A1").Value = "Order";
sheet.Cell("B1").Value = "Amount";
sheet.Cell("C1").Value = "Status";
sheet.Cell("A2").Value = "Order 1001";
sheet.Cell("B2").Value = 45.50;
sheet.Cell("C2").Value = "Pending";

var header = sheet.Range("A1:C1");
header.Style.Font.Bold = true;
header.Style.Fill.BackgroundColor = XlsxColor.LightBlue;
sheet.Range("B2:B20").Style.NumberFormat.Format = "#,##0.00";

sheet.Range("B2:B20").AddConditionalFormat()
    .WhenLessThan(50).Fill.SetBackgroundColor(XlsxColor.LightYellow);

sheet.Cell("E1").Value = "Pending";
sheet.Cell("E2").Value = "Paid";
sheet.Range("C2:C20").CreateDataValidation().List(sheet.Range("$E$1:$E$2"));
sheet.Column(5).Hide();

sheet.Cell("A2").CreateComment().AddText("Check the shipping address.");
sheet.Cell("A2").SetHyperlink(new XlsxHyperlink(new Uri("https://example.com/orders/1001")));
sheet.SheetView.FreezeRows(1);
sheet.Columns(1, 3).AdjustToContents();

workbook.SaveAs("orders.xlsx");
```

Conditional formatting determines how a spreadsheet application displays matching
cells. Data validation defines allowed input for users editing the workbook. Apply
the same checks in C# when your application writes values programmatically.
`CreateComment()` adds the legacy Excel note associated with a cell.
