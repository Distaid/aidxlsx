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
- a public API whose structure matches the upstream project after namespace normalization
  and the `XL*` → `Xlsx*` rename; source and binary compatibility with upstream is not
  offered;
- no `DocumentFormat.OpenXml` dependency at runtime: workbooks are read and written by a
  custom OPC/XML layer (`Aid.Xlsx.Opc`);
- preservation of unknown parts on round-trip, including `.xlsm` VBA projects;
- NativeAOT and trimming as supported scenarios, with reflection-free code on the common
  paths;
- a formula engine ported without semantic changes.

## Target framework and NativeAOT

Only `net10.0` is targeted — for the library, the tests and the samples alike. There is no
`netstandard`, no .NET Framework and no `net462` target.

NativeAOT is a product goal rather than an afterthought: the project plans to enable the
trim, AOT and single-file analyzers, to treat warnings as errors, and to run an AOT smoke
application through the create/save/load scenario on every supported platform in CI.

## Repository layout (planned)

```text
src/Aid.Xlsx/                   core library: domain folders at the top level
                                (Cells, Ranges, Styles, PivotTables, ...); part
                                readers and writers live in Internal/IO
src/Aid.Xlsx.Opc/               OPC container, relationships, XML/MCE primitives
src/Aid.Xlsx.Formula/           formula parser
src/Aid.Xlsx.SourceGenerators/  Roslyn incremental generators
src/Aid.Xlsx.CodeGen/           XSD-based code generation (build-time tool)
tests/                          unit, compatibility, AOT smoke and package checks
benchmarks/                     BenchmarkDotNet benchmarks
eng/                            build and release scripts
```

> The solution currently contains the imported library, the OPC/XML layer, the code
> generation tool and the test project. The formula parser, API compatibility checks,
> the NativeAOT smoke app and benchmarks are added in later phases.
>
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