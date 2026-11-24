# Troubleshooting

Most Aid.Xlsx failures are reported by a small set of library exceptions or by
standard .NET exceptions from the file and stream APIs. This guide maps symptoms to
causes, shows the fix for the common cases, explains the test suite and platform
gates, and lists what to include in a useful issue report.

## Exception types in the library

The public exception namespace contains three types. Verify the exact type before
catching it; several failures that look library-specific are standard exceptions.

| Type | Base type | Raised when |
| --- | --- | --- |
| `XlsxException` | `Exception` | Abstract base; not thrown directly. |
| `EmptyTableException` | `XlsxException` | Saving a table with no body rows. |
| `InvalidReferenceException` | `Exception` | A pivot cache source cannot be resolved. |

```csharp
using System;
using Aid.Xlsx.Model.Exceptions;
using Aid.Xlsx.Model.Workbooks;

try
{
    using var workbook = new XlsxWorkbook("input.xlsx");
    workbook.SaveAs("output.xlsx");
}
catch (EmptyTableException ex)
{
    Console.WriteLine($"A table has no rows: {ex.Message}");
}
catch (InvalidReferenceException ex)
{
    Console.WriteLine($"A pivot source is missing: {ex.Message}");
}
```

`InvalidReferenceException` is not derived from `XlsxException`; catch it separately.
`XlsxException` is abstract, so a `catch (XlsxException)` block catches library-derived
exceptions, including `EmptyTableException` and calculation-memory limit failures.
Standard exceptions are not wrapped in it.

The save-time `EmptyTableException` message is
`Table 'Items' should have at least 1 row.` A table created over a header-only range,
or a table whose data rows were deleted, hits this check. Add a data row or remove
the table before saving. `InvalidReferenceException` usually appears after a pivot
table's source sheet or range was deleted; the default message is
`Reference to the data is not valid.`

### Standard exceptions you will see

| Exception | Common trigger |
| --- | --- |
| `InvalidOperationException` | No worksheets; `Save()` on new; formula cycle; lock change. |
| `ArgumentException` | Bad extension, duplicate sheet/name, bad address, equal converter paths. |
| `ArgumentOutOfRangeException` | Worksheet position too large; negative calculation limit. |
| `KeyNotFoundException` | Missing worksheet, custom property, or collection name. |
| `InvalidDataException` | Not OPC/ZIP, missing content types, no workbook part, bad BIFF8. |
| `NotSupportedException` | Encrypted or pre-BIFF8 `.xls`; chart/macro/dialog sheets. |
| `IOException` | A file is locked or the share mode denies access. |
| `UnauthorizedAccessException` | Saving over a read-only file on Windows. |
| `ObjectDisposedException` | A workbook member is used after `Dispose`. |
| `NotImplementedException` | Array or data-table formula evaluation. |

## Common failures and fixes

### Save fails because there are no worksheets

Symptom: `InvalidOperationException: Workbooks need at least one worksheet.`

A workbook cannot be written with an empty `Worksheets` collection. Add a sheet
before the first save.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");
sheet.Cell("A1").Value = "Ready";

workbook.SaveAs("report.xlsx");
```

The same check also runs for `SaveAs`; there is no mode that writes a workbook without
a worksheet.

### Save fails because a new workbook has no source

Symptom:
`InvalidOperationException: This is a new file. Please use one of the 'SaveAs' methods.`

`Save()` writes back to the path or stream used by the constructor. A workbook created
with `new XlsxWorkbook()` has no source. Use `SaveAs(path)` or `SaveAs(stream)` for
the first write; later `Save()` calls work because `SaveAs` sets the new source.

```csharp
using var workbook = new XlsxWorkbook();
workbook.Worksheets.Add("Report");
workbook.SaveAs("report.xlsx");
workbook.Worksheet(1).Cell("A1").Value = "Second write";
workbook.Save();
```

### Invalid cell or range address

`Cell("A0")`, `Cell("1A")`, and `Range("A1:")` do not parse. The public members throw
`ArgumentException` with messages such as:

- `'1A' is not A1 address or workbook named range.`
- `'A1:' is not A1 address or named range.`

Use A1 notation, numeric rows and columns, or a defined name. A worksheet name in an
address must be quoted when it contains spaces, for example
`workbook.Range("'Monthly report'!A1:C3")`.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");

sheet.Cell("A1").Value = "A1";
sheet.Cell(2, 1).Value = "Row 2, column 1";
sheet.Range("A1:C3").Style.Font.Bold = true;
sheet.Range(4, 1, 4, 3).Style.Fill.BackgroundColor =
    Aid.Xlsx.Model.Styles.Colors.XlsxColor.LightYellow;
```

`workbook.CellFromFullAddress` and `workbook.RangeFromFullAddress` can resolve a full
address such as `Report!B2`; use them when the sheet name is part of the input string.

### Assigning Value removes a formula

`IXlsxCell.Value` is the value slot. Setting it clears `FormulaA1` on that cell, just
as typing a value into a formula cell in Excel does. Reading `Value` on a formula cell
evaluates the formula; reading `CachedValue` returns the stored result without
evaluation. Neither behavior is an error, which makes accidental formula loss easy to
miss.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
var sheet = workbook.Worksheet(1);
var cell = sheet.Cell("B10");

if (cell.HasFormula)
    Console.WriteLine(cell.FormulaA1); // Keep the formula for diagnostics.

cell.FormulaA1 = "SUM(B2:B9)";  // Writes or replaces a formula.
cell.Value = 42;                 // Removes the formula and writes a number.
```

To preserve formulas while editing other cells, do not clear or reassign formula
cells. To force a recalculation, call `workbook.RecalculateAllFormulas()` or set
`SaveOptions.EvaluateFormulasBeforeSaving = true`. A circular reference throws
`InvalidOperationException` with a message naming the cell, and an unsupported
array/data-table formula can throw `NotImplementedException`.

### File is locked or read-only

A path load closes its file handle when the workbook is constructed, so a loaded
workbook does not keep the file locked. Loading or saving still fails when another
process holds an exclusive lock.

```csharp
using System;
using System.IO;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("report.xlsx");

try
{
    workbook.Save();
}
catch (IOException ex)
{
    Console.WriteLine($"The file is in use: {ex.Message}");
}
catch (UnauthorizedAccessException ex)
{
    Console.WriteLine($"The file is read-only: {ex.Message}");
}
```

Close the spreadsheet application or save to a different path. On Windows, a file
with the `ReadOnly` attribute cannot be overwritten; the test suite verifies this and
excludes the check on Unix, where the read-only attribute does not block this code
path. If the source must stay untouched, load it and use `SaveAs` to write elsewhere.

### Stream is not seekable

Loading a non-seekable stream works for reading: the constructor copies it to an
internal `MemoryStream`. Later saves fail because the original stream is kept as the
save source: `SaveAs` throws `NotSupportedException` while rewinding it, and `Save()`
rejects it as not seekable. A loaded workbook also requires the stream passed to
`SaveAs(Stream)` to be readable, seekable, and writable.

Use a seekable source and a `MemoryStream` destination. For a web response, copy the
bytes after the save; never hand a write-only response body to `SaveAs`.

```csharp
using System.IO;
using Aid.Xlsx.Model.Workbooks;

using var source = new FileStream(
    "input.xlsx", FileMode.Open, FileAccess.Read, FileShare.Read);
using var output = new MemoryStream();

using (var workbook = new XlsxWorkbook(source))
{
    workbook.Worksheet(1).Cell("A1").Value = "Stream copy";
    workbook.SaveAs(output);
}

output.Position = 0;
using var response = File.Create("output.xlsx");
output.CopyTo(response);
```

The destination in the last lines is a seekable output stream. If the destination is
a response body, use a local `MemoryStream` for the save and then write
`output.ToArray()` to the response.

### Package is truncated or corrupt

Structural failures throw `InvalidDataException`. Common messages include:

- `The OPC container is missing [Content_Types].xml.`
- `The package has no workbook part.`
- `The workbook part does not contain a SpreadsheetML workbook.`
- `The OPC container has too many ZIP entries.`
- `The OPC container exceeds its uncompressed size limit.`
- `OPC entry '...' exceeds the compression ratio limit.`

```csharp
using System.IO;
using Aid.Xlsx.Model.Workbooks;

try
{
    using var workbook = new XlsxWorkbook("broken.xlsx");
}
catch (InvalidDataException ex)
{
    Console.WriteLine($"Not a readable workbook: {ex.Message}");
}
```

A truncated download or a ZIP produced by a non-OPC tool cannot be repaired by load
options. Re-export the file, verify its size, and check that the ZIP contains
`[Content_Types].xml` and a workbook part. Encrypted OOXML packages are outside the
supported scope; decrypt them in Excel first.

### Unsupported file format: `.xls` or an unknown extension

`XlsxWorkbook("legacy.xls")` fails before reading the content with
`ArgumentException: Extension 'xls' is not supported. Supported extensions are
'.xlsx', '.xlsm', '.xltx' and '.xltm'.` The path constructor determines the package
type from the extension, so a binary `.xls` file is not a valid source. Other
extensions such as `.xlsb` are rejected the same way; export them to `.xlsx` first.

Convert the workbook first. The converter accepts BIFF8 (Excel 97-2003) and writes a
new `.xlsx`; it does not modify the source.

```csharp
using Aid.Xlsx.Conversion;

XlsConverter.Convert("legacy.xls", "legacy.xlsx");
```

For an `.xls` stream, the ZIP layer throws `InvalidDataException` rather than the
extension message because there is no path to inspect. Encrypted `.xls` files throw
`NotSupportedException` (`Encrypted .xls workbooks are not supported.`); older BIFF
versions throw `NotSupportedException` as well.

### Culture-related parsing or formatting

The package parser reads XML numbers with invariant culture. Formula calculation uses
the current thread culture, and `GetFormattedString()` without an argument uses the
current culture. A server whose culture differs from the culture used to build the
workbook can therefore display different text or evaluate locale-sensitive formulas
differently.

```csharp
using System.Globalization;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
var sheet = workbook.Worksheet(1);

string german = sheet.Cell("B2").GetFormattedString(
    CultureInfo.GetCultureInfo("de-DE"));
string invariant = sheet.Cell("B2").GetFormattedString(
    CultureInfo.InvariantCulture);

Console.WriteLine($"{german} / {invariant}");
```

For deterministic behaviour, set `CultureInfo.CurrentCulture` explicitly at startup.
`XlsxWorkbook.EvaluateExpr` uses invariant culture; the instance `Evaluate` method and
the calculation engine use the current culture. The test suite pins the culture to
`en-US` for the whole assembly so results do not depend on the developer machine.

### Missing fonts and different column widths

Text measurement uses an embedded metric-compatible font for `Calibri`, so Calibri
column widths are the same on Windows and Linux. For a font that is not installed and
has no embedded substitute, the graphic engine falls back to its fallback font.
`AdjustToContents()` and conversions of stored column widths can then produce
different values on Linux.

This is a platform limitation, not a save bug. The suite excludes tests whose fixture
uses Arial or Yu Gothic on Unix for exactly this reason. If exact widths matter:

- install the font on the host, or
- supply the font bytes through the rendering API, or
- compare widths with a tolerance instead of exact equality.

`LoadOptions.GraphicEngine` and the static `LoadOptions.DefaultGraphicEngine` accept a
custom `IXlsxGraphicEngine`; `DefaultGraphicEngine.CreateOnlyWithFonts` builds one from
font streams when the file system is not available.

### Generated mapper is missing

`InsertMappedTable` or `InsertMappedData` throws
`NotSupportedException: Add [XlsxTable] to 'Order' and enable the Aid.Xlsx source
generator.` The row type is not annotated, the source generator is not referenced as
an analyzer, or the mapper did not compile. Add `[XlsxTable]` to the row type, confirm
the Aid.Xlsx package reference (it includes the generator), and fix any `XLSX001`,
`XLSX002`, `XLSX003`, or `XLSX005` compiler error. For a reflection-free alternative
without the generator, use `InsertTable(DataTable)` or `InsertData(DataTable)`.

## Run the test suite

From the repository root, restore once, then run the NUnit suite:

```powershell
dotnet test tests/Aid.Xlsx.Tests -c Release
```

Filter by namespace or test name while investigating:

```powershell
dotnet test tests/Aid.Xlsx.Tests -c Release `
  --filter "FullyQualifiedName~Aid.Xlsx.Tests.SpreadsheetML"
dotnet test tests/Aid.Xlsx.Tests -c Release `
  --filter "Name~CannotSaveAsOverwriteExistingReadOnlyFile"
```

The test project sets `en-US` as the assembly culture. If a test passes locally and
fails in CI, check the installed fonts, the file-system case sensitivity, and whether
the test is platform-gated.

### Platform-gated tests

Tests use NUnit's `[Platform(Exclude = "Unix")]` attribute when a platform cannot
provide the same primitive. Current gates cover:

| Area | Why the gate exists |
| --- | --- |
| Overwriting a read-only file | Windows rejects it; Unix truncate/rename succeeds. |
| Default column width with Arial | Arial is missing on Unix; fallback changes width. |
| Base column width with Arial | Same fallback issue for the stored base width. |
| Default width with Yu Gothic | The font is missing on Unix; the fallback changes the result. |

Do not loosen a metric assertion or remove a gate to make a platform pass. Gate only
when the platform genuinely lacks the font or primitive, and keep the reason next to
the attribute.

### Why Linux font metrics differ

`Calibri` is measured through the embedded metric-compatible font, so it is stable.
Other fonts are loaded from the operating system. A Linux container usually does not
have Arial, Yu Gothic, or many Microsoft Office fonts, so the engine falls back to the
embedded font. The glyph widths and therefore column widths differ. This affects
`AdjustToContents`, stored width conversion, and image placement; it does not corrupt
cells. See the [test project README](../tests/Aid.Xlsx.Tests/README.md) for the
fixture layout.

## File a useful issue

1. A minimal standalone program (usually 10-30 lines) that creates, loads, or saves a
   workbook and throws.
2. The exact Aid.Xlsx version from `dotnet list package` and the target framework.
3. `dotnet --info` output, including the OS, architecture, and SDK version.
4. The full exception type, message, and stack trace.
5. The `LoadOptions` and `SaveOptions` values in use, especially
   `RecalculateAllFormulas`, `EvaluateFormulasBeforeSaving`, and `ValidatePackage`.
6. Whether the file opens in Excel or LibreOffice; attach it if it is not confidential.
7. Whether the failure reproduces with `ValidatePackage = true`, and whether it occurs
   on Windows, Linux, or both.

For a CI-only failure, state the runtime identifier and whether the build used
`PublishTrimmed` or `PublishAot`. For a formula problem, include the formula text and
the expected result from Excel. Do not attach workbooks with credentials, personal
data, or macros you cannot share; reduce the file to the smallest failing sheet first.
