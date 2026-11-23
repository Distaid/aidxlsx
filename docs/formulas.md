# Formulas and recalculation

Assign an expression to `FormulaA1` for A1 references or to `FormulaR1C1` for R1C1
references. Use English function names and commas between function arguments. A
leading `=` is optional. Assigning `Value` to a formula cell replaces the formula.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Invoice");

sheet.Cell("A1").Value = 10;
sheet.Cell("A2").Value = 15;
sheet.Cell("A3").Value = 17;

var total = sheet.Cell("A4");
total.FormulaA1 = "SUM(A1:A3)";
sheet.Cell("B4").FormulaR1C1 = "RC[-1]*2";
workbook.RecalculateAllFormulas();

Console.WriteLine(total.GetDouble()); // 42
Console.WriteLine(sheet.Cell("B4").GetDouble()); // 84

workbook.SaveAs("invoice.xlsx", new SaveOptions
{
    EvaluateFormulasBeforeSaving = true
});
```

| API | Behavior |
| --- | --- |
| `FormulaA1` | Reads or sets the expression in A1 notation. |
| `CachedValue` | Reads the stored result without evaluating the formula. The result can be blank or stale. |
| `Value` | Reads the cell value. For a dirty formula, it can trigger evaluation before returning the result. |
| `RecalculateAllFormulas()` | Recalculates the workbook explicitly after edits. |
| `LoadOptions.RecalculateAllFormulas` | Requests a full recalculation when opening a workbook. |
| `SaveOptions.EvaluateFormulasBeforeSaving` | Evaluates formulas and writes available results to the XLSX package. By default, ordinary saves write formulas without their cached results. |

Check the functions used by your workbook before depending on calculation outside
Excel. If your application only needs to read the result already in a file, use
`CachedValue` and check whether it is present and appropriate for your workflow.
