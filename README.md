# Aid.Xlsx

**Status: pre-alpha.** The sources are imported and the solution builds, but this is an
early port onto an own OPC/XML layer with NativeAOT as a first-class goal. Nothing here is
a finished product, no package has been published yet, and the public API may change
without notice.

Aid.Xlsx is a .NET library for reading, creating and editing `.xlsx` and `.xlsm`
workbooks. It targets `net10.0`.

## Goals

- read, create and edit `.xlsx`/`.xlsm`: cells, values, formulas, cached values, styles,
  number formats, shared strings and rich text, tables, conditional formatting, data
  validation, comments, hyperlinks, images, page setup and protection;
- a public API organized by domain, with `Xlsx*` and `IXlsx*` type names;
- no `DocumentFormat.OpenXml` dependency at runtime: workbooks are read and written by a
  custom OPC/XML layer (`Aid.Xlsx.Opc`);
- preservation of unknown parts on round-trip, including `.xlsm` VBA projects;
- NativeAOT and trimming as supported scenarios, with reflection-free code on the common
  paths;
- a formula engine ported without semantic changes.

## Target framework and NativeAOT

The runtime library, tests and samples target `net10.0`. The build-time source generator
targets `netstandard2.0` so that it can run in the compiler host.

The runtime libraries enable trim, AOT and single-file analyzers. The smoke application
creates, saves and reloads a workbook after a fully trimmed or NativeAOT publish.

`[XlsxTable]` generates a mapper for public instance fields and readable properties. Use
`InsertMappedTable` or `InsertMappedData` to insert those rows without reflection. The NuGet
package includes the generator.

```csharp
using Aid.Xlsx.Attributes;
using Aid.Xlsx.InsertData;
using Aid.Xlsx.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");
sheet.Cell("A1").InsertMappedTable(new[] { new Sale("Pears", 22m) });
workbook.SaveAs("sales.xlsx");

[XlsxTable("SalesData")]
internal record Sale(string Item, [property: XlsxColumn(Format = "0.00")] decimal Amount);
```

## Convert an Excel 97–2003 workbook

```csharp
using Aid.Xlsx.Conversion;

XlsConverter.Convert("legacy.xls", "converted.xlsx");
```

The converter reads BIFF8 directly and also accepts streams. It carries over worksheet
values, cached formula results, number formats, row and column dimensions, visibility, and
merged cells. Formulas become static values. Charts, drawings, macros, encryption, and
formatting other than number formats are outside the supported conversion scope.

## Repository layout

```text
src/Aid.Xlsx/                   core library: domain folders at the top level
                                (Cells, Ranges, Styles, PivotTables, ...); part
                                readers and writers live in Internal/IO
src/Aid.Xlsx.Opc/               OPC container, relationships, XML/MCE primitives
src/Aid.Xlsx.Formula/           formula parser
src/Aid.Xlsx.Generators/        Roslyn incremental generators
src/Aid.Xlsx.CodeGen/           XSD-based code generation (build-time tool)
tests/                          unit, compatibility, AOT smoke and package checks
benchmarks/                     BenchmarkDotNet benchmarks
eng/                            build and release scripts
```

> The core uses a flat layout: domain folders sit directly under `src/Aid.Xlsx/`, while
> part readers and writers live in `src/Aid.Xlsx/Internal/IO/`.

## Building

The planned build toolchain pins the .NET SDK in `global.json` (`10.0.403`) and is started
from the repository root:

```powershell
dotnet build Aid.Xlsx.slnx -c Release
```

> `dotnet restore` needs network access on first run. The build is expected to be green
> for the imported projects; the remaining projects arrive in later phases.

## Relationship to ClosedXML

Aid.Xlsx is an independent project derived from
[ClosedXML](https://github.com/ClosedXML/ClosedXML) at commit
`057f5976c788d97e8eac86d1a528bd5411d9386e`. It is not affiliated with, endorsed by, or
sponsored by the ClosedXML project.

ClosedXML is distributed under the MIT License; the original copyright notice is
preserved in `LICENSE`. Attribution details and the list of third-party components are
in `NOTICE`.

## License

MIT, see [LICENSE](LICENSE). Attribution and third-party components are listed in
[NOTICE](NOTICE); third-party license texts live in `third-party/`.
