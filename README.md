# Aid.Xlsx

![Aid.Xlsx icon](https://raw.githubusercontent.com/Distaid/aidxlsx/main/assets/icon.png)

[![.NET 11](https://img.shields.io/badge/.NET-11.0-512BD4?logo=dotnet)](https://dotnet.microsoft.com/en-us/download/dotnet/11.0)
[![Latest GitHub release](https://img.shields.io/github/v/release/Distaid/aidxlsx?label=GitHub%20release)](https://github.com/Distaid/aidxlsx/releases/latest)
[![NuGet version](https://img.shields.io/nuget/v/Aid.Xlsx?logo=nuget&label=NuGet)](https://www.nuget.org/packages/Aid.Xlsx)
[![NuGet downloads](https://img.shields.io/nuget/dt/Aid.Xlsx?label=NuGet%20downloads)](https://www.nuget.org/packages/Aid.Xlsx)
[![CI status](https://img.shields.io/github/actions/workflow/status/Distaid/aidxlsx/ci.yml?branch=main&logo=githubactions&label=CI)](https://github.com/Distaid/aidxlsx/actions/workflows/ci.yml?query=branch%3Amain)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://github.com/Distaid/aidxlsx/blob/main/LICENSE)

Aid.Xlsx is a .NET library for reading, creating and editing `.xlsx` and `.xlsm`
workbooks. It targets `net11.0` and uses its own OPC/XML layer. The Open XML SDK is
used by the tests for schema validation.

Aid.Xlsx `2.x` targets .NET 11. For .NET 10 applications, use Aid.Xlsx `1.x`.

## Documentation

Create your first workbook with the [getting started guide](docs/getting-started.md).
Find recipes for cells, formatting, formulas, tables, streams, and `.xls` conversion
in the [documentation index](docs/README.md).

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

The runtime library, tests and samples target `net11.0`. The build-time source generator
targets `netstandard2.0` so that it can run in the compiler host.

The runtime libraries enable trim, AOT and single-file analyzers. The smoke application
creates, saves and reloads a workbook after a fully trimmed or NativeAOT publish.

`[XlsxTable]` generates a mapper for public instance fields and readable properties. Use
`InsertMappedTable` or `InsertMappedData` to insert those rows without reflection. The NuGet
package includes the generator.

```csharp
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.Workbooks;

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

XlsConverter.Convert("legacy.xls"); // Writes legacy.xlsx next to legacy.xls.
```

The converter reads BIFF8 directly. You can also specify a destination path or pass streams.
It carries over worksheet values, cell formatting, row and column dimensions, visibility, and
merged cells. It preserves formulas that use supported operators, references within the workbook,
and common worksheet functions, along with their cached results. Other formulas become static
values. See the compatibility guide for conversion limits.

## Repository layout

| Project | Responsibility |
| --- | --- |
| [Aid.Xlsx](src/Aid.Xlsx/README.md) | Workbook model, SpreadsheetML codecs, calculation, and BIFF8 conversion |
| [Aid.Xlsx.Opc](src/Aid.Xlsx.Opc/README.md) | Package storage, part relationships, and XML traversal |
| [Aid.Xlsx.Formula](src/Aid.Xlsx.Formula/README.md) | Formula and reference parsing |
| [Aid.Xlsx.Generators](src/Aid.Xlsx.Generators/README.md) | Compiler-host generators for table mapping and workbook metadata |
| [Aid.Xlsx.CodeGen](src/Aid.Xlsx.CodeGen/README.md) | Command-line generation of SpreadsheetML readers from the pinned XSD |

`tests/` contains the unit suite, compatibility checks, NativeAOT smoke test, fuzz targets, and
standalone benchmarks. Repository validation lives under `eng/`.

## Performance

Load and save timings, allocations, and commands to reproduce the measurements are in the
[workbook I/O benchmarks](docs/performance.md).

## Building

Use the .NET SDK pinned in `global.json` from the repository root.

```powershell
dotnet build Aid.Xlsx.slnx -c Release
```

The first restore needs network access. See the [migration guide](docs/migration-from-closedxml.md),
[compatibility notes](docs/compatibility.md), and [NativeAOT guide](docs/native-aot.md).

Run the GitHub `release` workflow on `main` with the version from `Directory.Build.props`.
This dry run uploads the NuGet and symbol packages, a CycloneDX dependency SBOM, and
SHA-256 checksums. The owner can then push the matching version tag on the verified commit
to publish using the `nuget-production` environment.

## License

Aid.Xlsx is an independent MIT-licensed project derived from ClosedXML. See [LICENSE](LICENSE)
and [NOTICE](NOTICE) for license and attribution details.
