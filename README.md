# Aid.Xlsx

**Status: pre-alpha.** No package has been published yet, and the public API may change.

Aid.Xlsx is a .NET library for reading, creating and editing `.xlsx` and `.xlsm`
workbooks. It targets `net10.0` and uses its own OPC/XML layer. The Open XML SDK is
used by the tests for schema validation.

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
values, cell formatting, row and column dimensions, visibility, and merged cells. It preserves
formulas that use supported operators, same-sheet references, and common worksheet functions,
along with their cached results. Other formulas become static values. See the compatibility
guide for conversion limits.

## Repository layout

```text
src/Aid.Xlsx/                   core library: domain folders at the top level
                                (Cells, Ranges, Styles, PivotTables, ...); part
                                readers and writers live in Internal/IO
src/Aid.Xlsx.Opc/               OPC container, relationships, XML/MCE primitives
src/Aid.Xlsx.Formula/           formula parser
src/Aid.Xlsx.Generators/        Roslyn incremental generators
src/Aid.Xlsx.CodeGen/           XSD-based code generation (build-time tool)
tests/                          unit, compatibility, AOT smoke, fuzzing and benchmarks
eng/                            build and release scripts
```

> The core uses a flat layout: domain folders sit directly under `src/Aid.Xlsx/`, while
> part readers and writers live in `src/Aid.Xlsx/Internal/IO/`.

## Building

Use the .NET SDK pinned in `global.json` (`10.0.403`) from the repository root.

```powershell
dotnet build Aid.Xlsx.slnx -c Release
```

The first restore needs network access. See the [migration guide](docs/migration-from-closedxml.md),
[compatibility notes](docs/compatibility.md), and [NativeAOT guide](docs/native-aot.md).

The GitHub release workflow can be dispatched to build the version in `Directory.Build.props`
as a dry run. The build uploads the NuGet and symbol packages, a CycloneDX dependency SBOM,
and SHA-256 checksums. Publication requires an owner-created version tag and the
`nuget-production` environment.

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
