# Formulas and recalculation

Assign a formula to `FormulaA1` or `FormulaR1C1`. A leading `=` is optional. Formulas use
English function names, commas between arguments, and a period as the decimal separator,
regardless of the user interface language. The library stores formulas in A1 form and
calculates them when you ask for a value or call a recalculation method.

## Write a formula

Use `FormulaA1` when references are written in A1 style. Use `FormulaR1C1` when a formula
is naturally relative to its own cell, such as a running total or a copied column formula.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Invoice");
sheet.Cell("A1").Value = 10;
sheet.Cell("A2").Value = 15;
sheet.Cell("A3").Value = 17;
sheet.Cell("A4").FormulaA1 = "SUM(A1:A3)";
sheet.Cell("B4").FormulaA1 = "=ROUND(A4/3, 2)"; // leading '=' is optional
sheet.Cell("C4").FormulaR1C1 = "RC[-1]*2";     // relative to C4

workbook.RecalculateAllFormulas();
Console.WriteLine(sheet.Cell("A4").GetDouble()); // 42
Console.WriteLine(sheet.Cell("B4").GetDouble()); // 14
Console.WriteLine(sheet.Cell("C4").GetDouble()); // 28
```

| Member | Behavior |
| --- | --- |
| `FormulaA1` | Gets or sets the formula in A1 notation; the getter has no `=`. |
| `FormulaR1C1` | Gets or sets the formula relative to the cell; the getter has no `=`. |
| `SetFormulaA1` / `SetFormulaR1C1` | Fluent forms that return the cell. |
| `HasFormula` | `true` when the cell has formula text. |
| `FormulaReference` | Range of an array formula; `null` for a normal formula. |

- `FormulaA1` trims whitespace, removes one leading `=`, and stores the text. Most text is
  validated only when it is evaluated; a malformed formula throws then, not necessarily at
  assignment.
- Function names are case-insensitive and English. Functions that Excel marks as future
  functions are stored with an `_xlfn.` prefix; write the plain name and let the library
  add the prefix. Reading `FormulaA1` may show the stored prefixed form.
- A semicolon is not an argument separator. Formula text copied from a localized formula
  bar that uses `;` fails to parse; replace semicolons with commas.
- Numbers inside formulas always use a period decimal separator, even when the current
  culture uses a comma.
- A formula assigned to an inferior cell of a merged range is ignored, like a value.
- Assigning an empty or whitespace string clears the formula.

## A1 versus R1C1

Use A1 for formulas that refer to fixed worksheet positions. Use R1C1 when the formula
should keep its relative shape, for example when it is generated in a loop or copied to
many cells.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Styles");
sheet.Cell("A1").Value = 2;
sheet.Cell("B1").FormulaA1 = "A1*10";
sheet.Cell("B2").FormulaR1C1 = "R[-1]C[-1]*10";
Console.WriteLine(sheet.Cell("B1").FormulaA1);   // A1*10
Console.WriteLine(sheet.Cell("B1").FormulaR1C1); // RC[-1]*10
Console.WriteLine(sheet.Cell("B2").FormulaA1);   // A1*10
Console.WriteLine(sheet.Cell("B2").FormulaR1C1); // R[-1]C[-1]*10
```

`$` fixes a row or column in A1 text. In R1C1, brackets are relative offsets and a bare
number is absolute (`R2C2` is `$B$2`). `FormulaR1C1` converts the expression to A1 at the
cell's position, so the same R1C1 text produces different A1 text in different cells.

`XlsxReferenceStyle` controls how addresses are displayed and which `refMode` is written to
the workbook. It does not change the formula syntax in the file, which is A1.

## Formula cells versus values

Use `CachedValue` to read the stored result without evaluating, and `Value` when you want
the current result. A cell holds either a value or a formula; assigning `Value` removes the
formula.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Exclusive");
var cell = sheet.Cell("A1");
cell.FormulaA1 = "1+1";
Console.WriteLine(cell.NeedsRecalculation);      // True
Console.WriteLine(cell.CachedValue.IsBlank);    // True

workbook.RecalculateAllFormulas();
Console.WriteLine(cell.CachedValue.GetNumber()); // 2
Console.WriteLine(cell.NeedsRecalculation);      // False
cell.Value = 5;                                  // replaces the formula
Console.WriteLine(cell.HasFormula);              // False
Console.WriteLine(cell.FormulaA1);               // empty
```

| Member | Evaluates? | Result |
| --- | --- | --- |
| `CachedValue` | No | Stored result, possibly `Blank` or stale. |
| `Value` | Yes, when the formula is dirty. | Current value after recalculation. |
| `DataType` | Yes, when the formula is dirty. | Type of the current value. |
| `NeedsRecalculation` | No | `true` when the formula must be evaluated. |
| `InvalidateFormula()` | No | Marks the formula dirty without changing it. |

- `Value` triggers a recalculation pass over dirty formulas in the workbook. A loop that
  reads many dirty formula cells can therefore be expensive; call
  `RecalculateAllFormulas()` once and then read.
- `CachedValue` is the stored result: the value from the file, or `Blank` when no result
  was saved (a new formula, or a formula written without evaluation).
- `FormulaA1` returns `string.Empty` for a value cell; it never returns `null`.
- `FormulaReference` is `null` for a normal formula cell and the covered block for a cell
  of an array formula.
- Setting a formula also sets the cell's string storage to inline so the result can be
  written next to the formula; setting a value restores the default.

## Array formulas

Use an array formula when one expression must fill a block, such as element-wise
multiplication of two columns. Set `FormulaArrayA1` on the whole target range.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Array");
sheet.Cell("A1").Value = 2;
sheet.Cell("A2").Value = 3;
sheet.Cell("A3").Value = 4;
sheet.Cell("B1").Value = 10;
sheet.Cell("B2").Value = 20;
sheet.Cell("B3").Value = 30;

sheet.Range("C1:C3").FormulaArrayA1 = "A1:A3*B1:B3";
workbook.RecalculateAllFormulas();
Console.WriteLine(sheet.Cell("C1").GetDouble());        // 20
Console.WriteLine(sheet.Cell("C2").GetDouble());        // 60
Console.WriteLine(sheet.Cell("C3").GetDouble());        // 120
Console.WriteLine(sheet.Cell("C2").HasArrayFormula);    // True
Console.WriteLine(sheet.Cell("C2").FormulaReference);    // C1:C3
```

- The assignment throws `InvalidOperationException` when the target intersects a merged
  range, intersects a table, or partially covers another array formula.
- Recalculation evaluates the expression from the master cell and broadcasts the result
  across the target. Extra cells receive `#N/A` when the result array is smaller.
- Only the master cell stores the formula text in the saved file; the other cells of the
  block are marked as part of the array. Excel displays the formula with braces.

## Recalculate and dirty state

Use explicit recalculation after a batch of edits, or when a workbook was loaded without
cached results. Formulas start dirty when created and when loaded without a stored value.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Recalc");
sheet.Cell("A1").Value = 4;
sheet.Cell("A2").FormulaA1 = "SQRT(A1)";
Console.WriteLine(sheet.Cell("A2").NeedsRecalculation); // True
workbook.RecalculateAllFormulas();
Console.WriteLine(sheet.Cell("A2").CachedValue.GetNumber()); // 2
sheet.Cell("A1").Value = 9;
Console.WriteLine(sheet.Cell("A2").NeedsRecalculation);      // True
sheet.RecalculateAllFormulas();                              // this sheet only
Console.WriteLine(sheet.Cell("A2").GetDouble());             // 3
```

| Call | Scope |
| --- | --- |
| `workbook.RecalculateAllFormulas()` | Marks every formula dirty, then evaluates them. |
| `worksheet.RecalculateAllFormulas()` | Marks and evaluates formulas on that worksheet only. |
| `LoadOptions.RecalculateAllFormulas` | Runs a workbook recalculation after loading. |
| `SaveOptions.EvaluateFormulasBeforeSaving` | Evaluates formulas while writing the package. |
| `XlsxCalculateMode` | `Auto`, `AutoNoTable`, `Manual`, `Default`; written to `calcPr`. |
| `CalculationOnSave`, `ForceFullCalculation` | Workbook flags saved to `calcPr`. |

- Worksheet-level recalculation does not evaluate formulas on other sheets. A dependent
  formula there stays dirty until its `Value` is read or the workbook is recalculated.
- `CalculateMode` and the related workbook flags are metadata for Excel. They do not change
  how this library evaluates formulas; use the two recalculation methods and save options.
- `LoadOptions.RecalculateAllFormulas` evaluates immediately after load. An unsupported
  formula makes the load throw instead of returning a workbook with a stale result.
- A circular reference throws `InvalidOperationException` from the calculation pass.
- Unsupported features such as external workbook references throw during evaluation; see
  the pitfalls below.

## Save formulas and cached results

Use `SaveOptions.EvaluateFormulasBeforeSaving` when another consumer must see results
without recalculating. The default save writes formula text without cached values.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Saved");
sheet.Cell("A1").Value = 6;
sheet.Cell("A2").FormulaA1 = "A1*7";
// Formula text is written, but no result is stored.
workbook.SaveAs("without-values.xlsx");
// Evaluate and write the result next to the formula.
workbook.SaveAs("with-values.xlsx", new SaveOptions
{
    EvaluateFormulasBeforeSaving = true
});

// Ask Excel to recalculate everything when it opens the file.
workbook.FullCalculationOnLoad = true;
workbook.SaveAs("recalc-on-open.xlsx");
```

| Setting | Effect |
| --- | --- |
| `EvaluateFormulasBeforeSaving` | Evaluates dirty formulas and writes results. Default `false`. |
| `GenerateCalculationChain` | Writes a calculation chain. Default `true`. |
| `FullCalculationOnLoad` | Sets `fullCalcOnLoad` so Excel recalculates on open. |
| `CalculationOnSave` | Sets `calcOnSave` in `calcPr`. |
| `ReferenceStyle` | Writes `A1` or `R1C1` as `refMode`; formula text stays A1. |

- With evaluation enabled, a formula that throws is written without a result; the save
  continues. This is why an unsupported function can leave a blank-looking formula cell.
- A non-dirty formula with a loaded cached value keeps that value when evaluation is
  enabled. A stale cached value is not refreshed unless the formula is marked dirty or the
  workbook is recalculated.
- Saving with `EvaluateFormulasBeforeSaving = false` discards the in-memory cached results.
  Reopening and reading `CachedValue` then returns `Blank`.
- `Save(bool validate, bool evaluateFormulae)` and the equivalent `SaveAs` overloads expose
  the same switch without constructing `SaveOptions`.

## Supported functions

The calculation engine registers about 217 functions. It is not a complete Excel
implementation; check the functions your workbook uses before relying on values computed
outside Excel, and see [Workbook compatibility](compatibility.md) for the supported surface.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Functions");
sheet.Cell("A1").Value = 10;
sheet.Cell("A2").Value = 20;
sheet.Cell("A3").FormulaA1 = "SUMIF(A1:A2, \">10\")";
sheet.Cell("A4").FormulaA1 = "TEXTJOIN(\"-\", TRUE, A1, A2)";
sheet.Cell("A5").FormulaA1 = "XLOOKUP(20, A1:A2, A1:A2)";
workbook.RecalculateAllFormulas();
Console.WriteLine(sheet.Cell("A3").GetDouble());   // 20
Console.WriteLine(sheet.Cell("A4").GetString());   // 10-20
Console.WriteLine(sheet.Cell("A5").GetError());    // NameNotRecognized
```

| Category | Count | Examples |
| --- | --- | --- |
| Math and trigonometry | 74 | `SUM`, `ROUND`, `SUMPRODUCT`, `SUBTOTAL`, `RAND`. |
| Statistical | 31 | `AVERAGE`, `COUNTIF`, `MEDIAN`, `STDEV.P`, `LARGE`. |
| Text | 31 | `CONCAT`, `TEXTJOIN`, `TEXT`, `VALUE`, `SUBSTITUTE`. |
| Date and time | 23 | `DATE`, `TODAY`, `NETWORKDAYS`, `WEEKNUM`, `YEARFRAC`. |
| Lookup and reference | 19 | `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `OFFSET`. |
| Information | 17 | `ISERROR`, `ISNUMBER`, `ISTEXT`, `N`, `NA`. |
| Database | 12 | `DSUM`, `DAVERAGE`, `DCOUNT`, `DGET`. |
| Logical | 7 | `IF`, `IFERROR`, `AND`, `OR`, `NOT`. |
| Financial | 3 | `FV`, `IPMT`, `PMT`. |

- Unsupported names evaluate to `#NAME?` rather than throwing. Common gaps include
  `XLOOKUP`, `IFS`, `SWITCH`, `LET`, `LAMBDA`, and dynamic-array functions such as
  `FILTER`, `SORT`, `UNIQUE`, and `SEQUENCE`.
- The Engineering category registers no functions. Only `FV`, `IPMT`, and `PMT` are
  available in Financial.
- Function names are resolved case-insensitively. Wrong argument counts for a known
  function throw `ExpressionParseException` when the formula is parsed.

## Errors, exceptions, and culture

Use `GetError()` for a formula that produced an Excel error, and catch
`ExpressionParseException` for text that is not a valid formula. Unknown functions produce
an error value, not an exception.

```csharp
using System;
using System.Globalization;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Errors");
sheet.Cell("A1").Value = 1;
sheet.Cell("A2").Value = 0;
sheet.Cell("A3").FormulaA1 = "A1/A2";
sheet.Cell("A5").FormulaA1 = "NOSUCHFUNCTION(1)";
workbook.RecalculateAllFormulas();
Console.WriteLine(sheet.Cell("A3").GetError());  // DivisionByZero
Console.WriteLine(sheet.Cell("A3").GetString()); // #DIV/0!
Console.WriteLine(sheet.Cell("A5").GetError());  // NameNotRecognized
CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
sheet.Cell("A6").FormulaA1 = "VALUE(\"1,5\")";
workbook.RecalculateAllFormulas();
Console.WriteLine(sheet.Cell("A6").GetDouble()); // 1.5

sheet.Cell("A4").FormulaA1 = "SUM(A1,";
try
{
    Console.WriteLine(sheet.Cell("A4").Value);
}
catch (ExpressionParseException exception)
{
    Console.WriteLine(exception.Message);
}

workbook.SaveAs("errors.xlsx");
```

| Failure | Result |
| --- | --- |
| Division by zero, bad reference, unknown name | `XlsxError` value; read it with `GetError()`. |
| Unknown function name | `#NAME?` value. |
| Malformed formula text | `ExpressionParseException` when evaluated. |
| Unsupported operator, 3D reference, external workbook reference | `NotImplementedException`. |

- Excel errors are values, not exceptions. `GetFormattedString()` displays the `#` code,
  and strict getters such as `GetDouble()` throw `InvalidCastException` on an error cell.
- Text-to-number coercion in functions such as `VALUE`, `TEXT`, and criteria arguments uses
  `CultureInfo.CurrentCulture`. The formula syntax itself stays invariant.
- `#REF!` is also what `FormulaReference` and range addresses render after the referenced
  range is deleted.

## Evaluate expressions outside cells

Use `Evaluate` for a quick check of an expression or a small calculation that does not
belong in a cell. Use `EvaluateExpr` for a culture-invariant literal expression with no
references.

```csharp
using System;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Evaluate");
sheet.Cell("A1").Value = 5;
sheet.Cell("A2").Value = 7;
XlsxCellValue sum = sheet.Evaluate("SUM(A1:A2)");
Console.WriteLine(sum.GetNumber()); // 12
XlsxCellValue power = XlsxWorkbook.EvaluateExpr("2^10");
Console.WriteLine(power.GetNumber()); // 1024

XlsxCellValue total = workbook.Evaluate("SUM(1, 2, 3)");
Console.WriteLine(total.GetNumber()); // 6

XlsxCellValue error = sheet.Evaluate("1/0");
Console.WriteLine(error.GetError()); // DivisionByZero
```

| Call | Context |
| --- | --- |
| `worksheet.Evaluate(expression)` | Has a worksheet, so ranges and sheet functions work. |
| `worksheet.Evaluate(expression, address)` | Adds a cell address for implicit intersection. |
| `workbook.Evaluate(expression)` | No worksheet; scalar functions work, range references do not. |
| `XlsxWorkbook.EvaluateExpr(expression)` | Static, invariant culture, no references. |

- `workbook.Evaluate("A1+A2")` throws `InvalidOperationException` because a reference
  needs a worksheet; use `worksheet.Evaluate` for anything that contains addresses.
- Results are `XlsxCellValue` values. Check `IsError` or call `GetError()` instead of
  assuming a number.

## Pitfalls

- **Cached values can be stale.** A value loaded from a file is not recomputed when a
  precedent changes. Call `RecalculateAllFormulas()` or read `Value` after edits.
- **Saving without evaluation erases results.** A plain `Save` writes formula text without
  a cached value; another reader sees `Blank` until it recalculates.
- **External references are not evaluated.** A formula that points into another workbook,
  or uses a 3D reference such as `Sheet1:Sheet3!A1`, throws `NotImplementedException`.
  The text is preserved, but this library cannot produce its value.
- **Data-table and shared formulas have no editing API.** Existing `t="dataTable"`
  formulas are written back but recalculating them throws `NotImplementedException`;
  shared formulas are normalized to one A1 formula per cell on save.
- **A failed calculation produces no cached result.** Saves with evaluation enabled skip
  the result for a formula that throws, leaving the cell empty in consumers that only read
  cached results.

For related material see [Cells and ranges](cells-and-ranges.md),
[Data types](data-types-and-conversions.md), [Compatibility](compatibility.md), and
[Workbook files](workbook-files.md).
