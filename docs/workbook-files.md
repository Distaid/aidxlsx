# Workbook files, streams, and templates

Aid.Xlsx reads and writes workbooks as OPC/ZIP packages: workbook XML, worksheet
parts, styles, shared strings, tables, drawings, and metadata. This guide covers
loading, saving, streams, templates, metadata, defined names, worksheet management,
and `.xls` conversion.

## Open a workbook from a path

The path constructor loads an existing `.xlsx`, `.xlsm`, `.xltx`, or `.xltm` package
before it returns. The source handle is closed when loading finishes, so the workbook
does not keep the file locked.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
var report = workbook.Worksheet("Report");

Console.WriteLine(report.Cell("B2").GetFormattedString());
```

`Worksheet(string)` throws `KeyNotFoundException` when the name is absent.
`Worksheet(int)` uses a 1-based position and throws `ArgumentException` when the
position is missing or ambiguous. `TryGetWorksheet` avoids the exception; the
worksheet-management section below shows it.

Edge cases:

- `.xls` is rejected before loading; convert it with `XlsConverter`.
- A non-OPC file throws `InvalidDataException`. Encrypted OOXML is not supported.
- A file opened exclusively by another process throws `IOException`.
- Relative paths resolve against the process working directory.

## Control loading with LoadOptions

These are all public load options. There is no shared-string toggle and no culture
property.

| Option | Default | Effect |
| --- | --- | --- |
| `RecalculateAllFormulas` | `false` | Recalculates every formula after loading. |
| `StrictAttributeParsing` | `true` | Throws on attributes with an invalid simple type. |
| `GraphicEngine` | `null` | Text and image metrics for this workbook. |
| `DefaultGraphicEngine` | `null` | Static fallback when `GraphicEngine` is null. |
| `Dpi` | `96, 96` | Image sizing and text measurement; must be positive. |
| `MaxCalculationMemoryBytes` | `0` | Calculation buffer limit; `0` means unlimited. |
| `CancellationToken` | `default` | Cancels recalculation started by later calls. |

```csharp
using Aid.Xlsx.Model.Workbooks;

var options = new LoadOptions
{
    RecalculateAllFormulas = true,
    StrictAttributeParsing = false,
    Dpi = new System.Drawing.Point(120, 120),
    MaxCalculationMemoryBytes = 256L * 1024 * 1024
};

using var workbook = new XlsxWorkbook("input.xlsx", options);
```

- Disabling `StrictAttributeParsing` treats malformed optional attributes as absent.
  It hides producer defects; enable it while investigating a file.
- Recalculation can be expensive and can throw on a bad formula. Enable it only when
  cached values are known to be stale.
- The calculation engine uses the current thread culture. Set
  `CultureInfo.CurrentCulture` explicitly for deterministic evaluation; XML parsing
  itself is culture-invariant.

## Read-only and editable streams

The stream constructors accept a readable stream. A seekable stream is read from its
current position; position it at the start. A non-seekable stream is copied to an
internal `MemoryStream`. The workbook does not dispose caller-owned streams. Open the
source with `FileAccess.ReadWrite` when the same stream must be saved with `Save()`.

```csharp
using System.IO;
using Aid.Xlsx.Model.Workbooks;

using var input = new FileStream("input.xlsx", FileMode.Open, FileAccess.Read);
using var output = new MemoryStream();

using (var workbook = new XlsxWorkbook(input))
{
    workbook.Worksheet(1).Cell("A1").Value = "Copied through a stream";
    workbook.SaveAs(output);
}

output.Position = 0;
using var file = File.Create("output.xlsx");
output.CopyTo(file);
```

`SaveAs(Stream)` on a loaded workbook copies the original package into the destination
and rewrites it there. The destination must be readable, seekable, and writable. A
write-only response body does not qualify; use the
[ASP.NET recipe](cookbook.md#stream-a-workbook-from-aspnet).

Edge cases:

- A non-seekable input can be read, but later saves fail. `SaveAs` throws
  `NotSupportedException` while rewinding; `Save` rejects the stream as not seekable.
- `Save()` writes to the original stream at position zero and calls `SetLength`.
  The source must be writable and seekable.
- `SaveAs(stream)` makes that stream the source; keep it open until the workbook is disposed.

## Save and SaveAs

| Call | Writes to |
| --- | --- |
| `Save()` | Original path or stream; throws for a new workbook. |
| `Save(SaveOptions)` | Same, with explicit options. |
| `SaveAs(path)` | A path; creates folders and replaces the file. |
| `SaveAs(path, SaveOptions)` | Same, with explicit options. |
| `SaveAs(stream)` | A stream; rewrites the package and changes the source. |
| `SaveAs(stream, SaveOptions)` | Same, with explicit options. |

The `bool` overloads pass `ValidatePackage` and `EvaluateFormulasBeforeSaving`.

| Option | Default | Effect |
| --- | --- | --- |
| `EvaluateFormulasBeforeSaving` | `false` | Evaluates formulas and stores available results. |
| `ValidatePackage` | `false`/`true` | Checks relationships and supported XML. |
| `GenerateCalculationChain` | `true` | Writes a calculation chain when formulas exist. |
| `ConsolidateConditionalFormatRanges` | `true` | Merges adjacent rule ranges. |
| `ConsolidateDataValidationRanges` | `true` | Merges adjacent validation ranges. |
| `FilterPrivacy` | `null` | Writes the `filterPrivacy` flag; `null` leaves it. |

`ValidatePackage` defaults to `true` in Debug and `false` in Release builds.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
workbook.Worksheet(1).Cell("B2").FormulaA1 = "SUM(B3:B12)";

workbook.Save(new SaveOptions
{
    EvaluateFormulasBeforeSaving = true,
    GenerateCalculationChain = true,
    ValidatePackage = true
});
```

### Overwrite behavior and atomicity

`SaveAs(path)` overwrites the target. A loaded file copied to a new target is first
copied and then rewritten; a new workbook deletes an existing target before creating
the package. The rewrite of an existing package is buffered in a `MemoryStream`
before the destination is truncated, but the operation is not a transactional file
replace. A failure during the final copy or a new-file save can leave a partial or
missing target.

For durable output, save to a temporary path and move it into place; prefer `MemoryStream`.
A new workbook can target a write-only stream through an internal buffer, but that path is
single-use.

The extension is case-insensitive and selects the content type: `.xlsx` workbook,
`.xlsm` macro-enabled workbook, `.xltx` template, `.xltm` macro-enabled template. It
does not convert macros.

## Templates and worksheet copies

`XlsxWorkbook.OpenFromTemplate(path)` loads a workbook as new: relationship IDs are
cleared, the source is never modified, and a save target is required.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = XlsxWorkbook.OpenFromTemplate("monthly-template.xlsx");
var sheet = workbook.Worksheet("Report");
sheet.Cell("B3").Value = 4200;
workbook.SaveAs("monthly-2026-01.xlsx");
```

Use `CopyTo` to duplicate a worksheet in the same or another workbook. Values, styles,
dimensions, merges, tables, pictures, validations, conditional formats, page setup,
and view state are copied.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
var source = workbook.Worksheet("Report");
var copy = source.CopyTo("Report copy", 2);
copy.Cell("A1").Value = "Copy of the report";
workbook.Save();
```

- A copied sheet needs a unique name; a duplicate throws `ArgumentException`.
- `CopyTo(workbook, newName)` copies into another workbook.
- `CopyTo` on a deleted sheet throws `InvalidOperationException`.
- `OpenFromTemplate` keeps the source marked as new, so `Save()` throws; use
  `SaveAs`. Keep `.xlsm` or `.xltm` when the template's VBA project must stay active.

## File properties and custom properties

Use `workbook.Properties` (an `XlsxWorkbookProperties`) for file metadata. The
workbook-level `Author` property is a model string initialized from the OS user name;
it is not written to the core properties. Set `Properties.Author` when the file author
matters.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
workbook.Properties.Title = "Monthly sales";
workbook.Properties.Author = "Finance";
workbook.Properties.Company = "Example Ltd";

workbook.CustomProperties.Add("Reviewed", true);
workbook.SaveAs("input-metadata.xlsx");
```

The available properties are `Author`, `Title`, `Subject`, `Category`, `Keywords`,
`Comments`, `Status`, `Created`, `Modified`, `LastModifiedBy`, `Company`, and
`Manager`.

- Custom values should be `string`, `DateTime`, numeric, or `bool`. Other object
  types fail when the package is written.
- A string that parses as a number is stored as a number.
- Duplicate custom names throw `ArgumentException`; missing names throw
  `KeyNotFoundException`.

## Defined names

Workbook-scoped names are visible from every sheet; sheet-scoped names are resolved on
their own sheet first. Use the matching collection to create each kind.

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

workbook.SaveAs("invoice.xlsx", new SaveOptions
{
    EvaluateFormulasBeforeSaving = true
});
```

- A name must start with a letter, an underscore, or a backslash and is limited to
  255 characters. Duplicate names throw `ArgumentException`.
- Use the range overload for workbook-scoped names; it writes the sheet qualifier.
- Assigning `RefersTo` directly rejects local A1 references without a sheet; prefer
  the range overloads.
- Deleting a name does not delete cells. A name invalidated by a deleted sheet stays
  in the workbook with `IsValid == false`.

## Calculation mode and reference style

The calculation settings are written to `calcPr` and control whether another
spreadsheet application recalculates the saved file.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
workbook.CalculateMode = XlsxCalculateMode.Auto;
workbook.ReferenceStyle = XlsxReferenceStyle.A1;
workbook.CalculationOnSave = true;
workbook.FullCalculationOnLoad = true;
workbook.Save();
```

`Default` omits the mode or style. `A1` and `R1C1` write `refMode`, but formulas in the
package remain A1 formulas; use `FormulaA1` or `FormulaR1C1` for formula syntax.
Clearing the flags can change whether Excel recalculates the file.

## Add, copy, move, and delete worksheets

`Worksheets` is an ordered 1-based collection. `Add(name)` appends;
`Add(name, position)` inserts. `Delete(name)`, `Delete(position)`, and
`worksheet.Delete()` remove sheets. Setting `Position` moves a sheet.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
workbook.Worksheets.Add("Summary", 1);
var detail = workbook.Worksheets.Add("Detail");

var copy = detail.CopyTo("Detail copy", 3);
copy.Position = 2;

if (workbook.Worksheets.TryGetWorksheet("Detail", out var found))
    found.Cell("A2").Value = "Found by name";

workbook.Worksheets.Delete("Summary");
workbook.SaveAs("reorganized.xlsx");
```

- `Worksheet("missing")` throws `KeyNotFoundException`; a missing or ambiguous
  position throws `ArgumentException`.
- Setting `Position` accepts `1` through `worksheet count + 1`; larger values throw
  `ArgumentOutOfRangeException`.
- Deleting the last worksheet leaves a workbook that cannot be saved.
- Unsupported sheet types such as chart sheets are preserved as package data and are
  not exposed through `Worksheets`.
- The workbook model exposes no public load, save, or change events.

## Convert a legacy `.xls` workbook

`XlsConverter` reads Excel 97-2003 BIFF8 workbooks. It copies values, supported
formulas with cached results, formatting, dimensions, visibility, and merged ranges.
Other formulas become static values.

```csharp
using System;
using System.IO;
using Aid.Xlsx.Conversion;

try
{
    XlsConverter.Convert("legacy.xls");
    XlsConverter.Convert("legacy.xls", "converted.xlsx");
}
catch (NotSupportedException ex)
{
    Console.WriteLine($"Unsupported: {ex.Message}");
}
catch (InvalidDataException ex)
{
    Console.WriteLine($"Malformed: {ex.Message}");
}
```

| Overload | Behavior |
| --- | --- |
| `Convert(sourcePath)` | Writes `.xlsx` next to the source. |
| `Convert(sourcePath, destinationPath)` | Writes to a different path. |
| `Convert(source, destination)` | Reads a seekable source; writes a new stream. |

- Only BIFF8 is accepted. Older BIFF versions and encrypted files throw
  `NotSupportedException`.
- Chart, macro, and dialog sheets throw `NotSupportedException`.
- Truncated or invalid records throw `InvalidDataException`.
- The stream overload needs a readable, seekable source and a different writable
  destination; both streams stay open.
- Named styles, rich text, conditional formatting, images, charts, and macros are not
  converted. See [Workbook compatibility](compatibility.md) for the formula subset.

## What a round trip preserves

Preserved:

- Unknown package parts, including VBA projects, custom binary parts, disconnected
  parts, and unsupported DrawingML shapes. Tests compare the VBA project and a custom
  part byte-for-byte.
- Values, formulas, styles, shared strings, tables, validations, conditional
  formats, comments, hyperlinks, images, page setup, protection, defined names, and
  visibility for modeled features.
- The content type when the output extension is unchanged.

Not preserved or not modeled:

- Charts have no public API. They are package data; editing a modeled drawing part
  can rewrite markup the object model does not represent.
- VBA is opaque: bytes survive, but macros cannot be inspected or edited. Keep
  `.xlsm` or `.xltm` to keep the macro-enabled content type.
- Cached formula results are not written by default. Set
  `EvaluateFormulasBeforeSaving` to keep values available without recalculation.
- Unsupported sheet types and package extensions are preserved as parts, not modeled.

## Disposal order

Dispose the workbook after the final save. Streams stay caller-owned.

```csharp
using System.IO;
using Aid.Xlsx.Model.Workbooks;

using var output = new MemoryStream();
using (var workbook = new XlsxWorkbook("input.xlsx"))
{
    workbook.Worksheet(1).Cell("A1").Value = "Ready";
    workbook.SaveAs(output);
}
output.Position = 0;
```

`Dispose` releases the model and worksheet caches. It does not save, close `output`,
or rewind the stream. After `SaveAs` returns, the bytes are complete and can be read
once the stream is rewound.
