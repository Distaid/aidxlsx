# Getting started

Aid.Xlsx creates, reads, and edits `.xlsx` and `.xlsm` workbooks from .NET. The
2.x line targets `net11.0`; the 1.x line targets `net10.0`. This guide installs the
package, creates a workbook, reopens it, and saves an edited copy.

## Choose the package version

| Application target | Package and SDK | Notes |
| --- | --- | --- |
| `net11.0` | Aid.Xlsx 2.x; .NET SDK 11.0.100+ | Current line; AOT and trimming supported. |
| `net10.0` | Aid.Xlsx 1.x; .NET SDK 10.0.100+ | Use while the application stays on .NET 10. |

Do not mix 1.x and 2.x packages in one application. The 2.x package does not install
into a `net10.0` project, and 1.x is not the supported line for `net11.0` projects.

## Install the package

### From NuGet

Create a console application and add the package.

```sh
dotnet new console --framework net11.0 --name WorkbookExample
cd WorkbookExample
dotnet add package Aid.Xlsx --version 2.0.1
```

For a .NET 10 application, use the 1.x package instead.

```sh
dotnet new console --framework net10.0 --name WorkbookExample
cd WorkbookExample
dotnet add package Aid.Xlsx --version 1.0.0
```

The repository pins SDK `11.0.100` in `global.json`. A consumer application can use
any .NET 11 SDK feature band; verify the installed SDK with `dotnet --list-sdks`.

### Pack the library from source

From the repository root, pack the runtime package. The package includes the table
mapper source generator.

```sh
dotnet pack src/Aid.Xlsx/Aid.Xlsx.csproj -c Release -o artifacts/packages
```

From the application directory, add the package without restoring, then restore from
the local package folder and NuGet.org. Replace `2.0.1` with the version in the
generated `.nupkg` filename if it differs.

```powershell
dotnet add package Aid.Xlsx --version 2.0.1 --no-restore
dotnet restore `
  --source "D:\Projects\AidXlsx\artifacts\packages" `
  --source https://api.nuget.org/v3/index.json
```

Package restore needs network access the first time because the runtime dependencies
are resolved from NuGet.org. After restore, the project can build offline.

## Create a workbook

Replace `Program.cs` with the example below. A workbook must contain at least one
worksheet before it can be saved. The `using` declaration disposes the workbook at the
end of the program.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");

sheet.Cell("A1").Value = "Item";
sheet.Cell("B1").Value = "Units";
sheet.Cell("C1").Value = "Price";
sheet.Cell("A2").Value = "Pears";
sheet.Cell(2, 2).Value = 12;
sheet.Cell("C2").Value = 1.25m;
sheet.Cell("B4").FormulaA1 = "B2*2";

workbook.SaveAs("sales.xlsx");
Console.WriteLine("Created sales.xlsx.");
```

Run it from the application directory:

```sh
dotnet run
```

`SaveAs` creates the file relative to the current working directory when the path is
not absolute. The file name extension selects the package type: `.xlsx`, `.xlsm`,
`.xltx`, or `.xltm`. Any other extension is rejected with `ArgumentException`.

## Reopen, read, and edit a copy

Replace `Program.cs` with this example and run it from the same directory. It prints
`12`, changes the quantity, and writes a new file. The original `sales.xlsx` stays
unchanged because the save target is different.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("sales.xlsx");
var sheet = workbook.Worksheet("Sales");

Console.WriteLine(sheet.Cell("B2").GetDouble()); // 12
sheet.Cell("B2").Value = 15;
sheet.Cell("A4").Value = "Updated";

workbook.SaveAs("sales-updated.xlsx");
```

Use `GetText()` for text, `GetDouble()` for numbers, and `GetDateTime()` for dates.
`TryGetValue<T>()` returns `false` instead of throwing when the stored value cannot be
converted. `GetFormattedString()` applies the cell's number format.

Saving to the same path uses `Save()`:

```csharp
using var workbook = new XlsxWorkbook("sales.xlsx");
workbook.Worksheet(1).Cell("B2").Value = 20;
workbook.Save();
```

`Save()` requires an existing file or writable stream source. A new workbook has no
source, so a new workbook must be written with `SaveAs(...)`.

## Dispose correctly

`XlsxWorkbook` implements `IDisposable`. Dispose it after the last read or write. A
`using` declaration is enough in a console application; the following form makes the
ordering explicit.

```csharp
using Aid.Xlsx.Model.Workbooks;

var workbook = new XlsxWorkbook("sales.xlsx");
try
{
    workbook.Worksheet(1).Cell("B2").Value = 20;
    workbook.Save();
}
finally
{
    workbook.Dispose();
}
```

Disposal releases the workbook model. It does not save and does not dispose streams
that the caller supplied. Using a disposed workbook throws `ObjectDisposedException`.
Keep caller-owned streams alive until the workbook is disposed and every save is
finished.

## Guided tour

| Guide | Read it for |
| --- | --- |
| [Cells and ranges](cells-and-ranges.md) | A1 addresses, typed reads, used cells, merges. |
| [Formatting](formatting-and-worksheets.md) | Styles, formats, rules, validation, notes. |
| [Formulas and recalculation](formulas.md) | A1/R1C1 formulas, cached values, recalculation. |
| [Tables](tables-and-data-import.md) | Tables, filters, `DataTable` import, row mappers. |
| [Workbook files](workbook-files.md) | Load/save options, streams, templates, names. |
| [NativeAOT and trimming](native-aot.md) | Reflection-free imports, warnings, publish matrix. |
| [Troubleshooting](troubleshooting.md) | Exception types, common failures, tests, issues. |
| [Cookbook](cookbook.md) | Recipes for reports, tables, images, web APIs. |
| [Workbook compatibility](compatibility.md) | Formats, preserved parts, conversion limits. |
| [Migration guide](migration-from-closedxml.md) | Renamed types and replacement APIs. |
| [Performance](performance.md) | Benchmarks and reproduction commands. |

### A practical first path

1. Finish the create, reopen, and dispose examples on this page.
2. Read [Cells and ranges](cells-and-ranges.md) to read values without guessing their
   type.
3. Read [Workbook files and streams](workbook-files.md) before loading production files
   or writing to a web response.
4. Use [Tables and data import](tables-and-data-import.md) plus
   [NativeAOT and trimming](native-aot.md) when the application imports records or is
   published with `PublishTrimmed` or `PublishAot`.

## Common first-run mistakes

- **`InvalidOperationException: Workbooks need at least one worksheet.`** No worksheet
  was added before save. Call `workbook.Worksheets.Add(...)` first.
- **`InvalidOperationException: This is a new file. Please use one of the 'SaveAs'
  methods.`** `Save()` needs an existing path or stream. A new workbook must start
  with `SaveAs`.
- **`NU1202` or a target-framework error.** Aid.Xlsx 2.x is installed into a
  `net10.0` project. Target `net11.0`, or install Aid.Xlsx 1.x.
- **`ArgumentException: Extension 'xls' is not supported.`** A legacy binary
  workbook was passed to `SaveAs`. Convert it with `XlsConverter` first.
- **The file is missing after the program exits.** Disposal does not save. Call
  `Save` or `SaveAs` while the workbook is alive.
- **`ObjectDisposedException`.** A workbook member was used after the `using` scope.
  Move the work inside the scope or dispose later.
- **Local package restore cannot find `2.0.1`.** The pack output or version does not
  match. Pack from the repository root and use the `.nupkg` file name.

## What to read next

If the application only creates a workbook, continue with
[Cells and ranges](cells-and-ranges.md) and
[Formatting and worksheet features](formatting-and-worksheets.md). If it loads files
produced by other tools, read [Workbook files and streams](workbook-files.md) and
[Troubleshooting](troubleshooting.md) first. For a complete end-to-end sample, use the
[Cookbook](cookbook.md).
