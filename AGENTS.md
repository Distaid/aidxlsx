# Agent guidance

This file is the entry point for anyone (human or AI) making changes in this repository.

## What this project is

`Aid.Xlsx` is a .NET library for reading, creating and editing `.xlsx` and `.xlsm`
workbooks. It targets `net10.0` and is built around its own OPC/XML layer instead of the
OpenXml SDK, with NativeAOT as a first-class goal.

## Prerequisites

- The .NET SDK version pinned in `global.json` is required and used as-is
  (`rollForward: disable`). Do not float the SDK: the lock files depend on it.
- After the first restore the build works offline.
- On Windows keep `core.autocrlf=false` in this clone; line endings are controlled by
  `.gitattributes` (`eol=lf` for everything except `*.ps1`, which is `crlf`).

## Build, test, pack

```powershell
dotnet restore Aid.Xlsx.slnx
dotnet build Aid.Xlsx.slnx -c Release -warnaserror
dotnet test Aid.Xlsx.slnx -c Release
dotnet pack src/Aid.Xlsx/Aid.Xlsx.csproj -c Release -o artifacts/packages
```

Repository checks that run in CI and are worth running locally:

```powershell
dotnet format Aid.Xlsx.slnx --verify-no-changes
powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\validate-workflows.ps1
```

`eng/scripts/validate-workflows.ps1` verifies the workflow files, encodings, package
metadata and the repository conventions described below.

## Project layout

| Path | Purpose |
|---|---|
| `src/Aid.Xlsx` | The library: cell model, ranges, styles, tables, drawings, CalcEngine, part readers and writers (`Internal/IO`) |
| `src/Aid.Xlsx.Opc` | OPC container, relationships, content types, XML/MCE primitives. Knows nothing about the Excel model |
| `src/Aid.Xlsx.CodeGen` | Build-time generator: reads OOXML schemas, emits the `*.g.cs` readers |
| `tests/Aid.Xlsx.Tests` | NUnit test suite |
| `eng/scripts` | Import and validation scripts |
| `.github/workflows` | CI gates |

The library uses a flat domain layout: domain folders sit directly under `src/Aid.Xlsx/`,
the only nested convenience folder is `Internal/IO`.

## Code conventions

- **Language:** code, comments, XML documentation, CLI messages, workflow text and
  documentation are written in English. Test string literals keep their original language
  when the test data itself is not English (text parsing, Unicode fixtures) — never
  "translate" test data.
- **Encoding:** UTF-8 without BOM for every file, except `*.ps1`/`*.psm1`, which must have
  a BOM so Windows PowerShell reads them correctly.
- **Warnings are errors.** `TreatWarningsAsErrors` is enabled for the whole solution.
  Do not add `NoWarn`, `#pragma warning disable`, `UnconditionalSuppressMessage` or
  `WarningsNotAsErrors` to make a warning disappear; fix the cause. Exceptions need a
  comment, a reason and a tracking note.
- **AOT and trimming are product requirements.** The library projects set
  `IsTrimmable`/`IsAotCompatible` and the trim/AOT analyzers are enabled. Code that needs
  reflection must be marked with `[RequiresUnreferencedCode]`/`[RequiresDynamicCode]`
  and must have a reflection-free path. Do not introduce `dynamic`, `Expression.Compile`,
  `Activator.CreateInstance(Type)`, `Array.CreateInstance(Type, ...)`,
  `Enum.GetValues(Type)` or `Type.GetType(string)` on common paths.
- **Public API:** Keep the `Xlsx*`/`IXlsx*` names and use the namespace for the owning
  domain. Keep future public API changes intentional and do not add legacy aliases.
- **Fonts:** metrics of `Calibri` always come from the embedded metric-compatible font,
  so column widths are identical on every OS. Do not loosen metric assertions to make a
  platform pass; if a font is genuinely missing on a platform, gate the test explicitly
  with `[Platform(Exclude = "...")]` and a comment.
- **The test project disables nullable analysis** on purpose: the tests come from a code
  base that predates nullable reference types. The library projects keep it enabled.
  Do not make the tests project nullable-enabled indirectly (for example with `#nullable
  enable` in a single file) without a reason.
- **C# file names:** Use one dot, before the `.cs` extension, including for partial classes.
  Generated sources use the `.g.cs` suffix.
- **Generated code:** `*.g.cs` files are produced by `Aid.Xlsx.CodeGen`
  (`dotnet run --project src/Aid.Xlsx.CodeGen -- all`) and committed. Never edit them by
  hand; change the generator or its inputs instead.
- **Lock files:** `packages.lock.json` files are committed. Do not edit them by hand;
  restore regenerates them. When changing dependencies, regenerate the lock files and
  the package versions in `Directory.Packages.props` in the same commit.

## Repository hygiene

- Local planning notes and research material are not part of this repository. Committed
  files must not reference them.
- The upstream project this library derives from is credited in `NOTICE`, `LICENSE` and
  `README.md`. Keep that attribution in those files and keep it out of commit messages,
  pull request titles and code comments.
- `LICENSE`, `NOTICE` and third-party license texts must stay in the package.
- Keep `CHANGELOG.md` up to date under `Unreleased`; do not invent release dates.

## Commit convention

Commits use Conventional Commits with a **scope**. Format:

```text
type(scope): subject

body (why the change is needed, what changed, tradeoffs)

footer (breaking changes, references)
```

Rules:

- `type` is one of `feat`, `fix`, `build`, `ci`, `docs`, `test`, `refactor`, `perf`,
  `chore`, `revert`.
- `scope` is one of the areas below. Pick the narrowest scope that owns the change; if a
  commit really spans areas, split it.
- `subject` is imperative, lowercase, at most 72 characters, no trailing period, no
  project-name mentions.
- A body is expected for anything non-trivial: explain the reason, not the diff.
- One logical change per commit. Never mix a formatting-only change with a behavior
  change; formatting passes go into their own `style`-scoped commit.
- Do not put internal planning or task identifiers into messages.

| Scope | Covers |
|---|---|
| `core` | `src/Aid.Xlsx` model, ranges, cells, worksheets, workbook |
| `opc` | `src/Aid.Xlsx.Opc` container, relationships, content types, XML primitives |
| `io` | part readers and writers under `src/Aid.Xlsx/Internal/IO` |
| `calc` | CalcEngine and functions |
| `styles` | styles, number formats, colors |
| `tables` | tables, filters, structured references |
| `drawings` | drawings, pictures, charts, comments, VML |
| `codegen` | schema-driven generator and generated `*.g.cs` |
| `tests` | test suite and fixtures |
| `build` | `Directory.Build.props`, `Directory.Packages.props`, `global.json`, SDK and analyzer settings |
| `packaging` | nuspec/package contents, symbols, icon, license files in the package |
| `deps` | dependency versions and lock files |
| `ci` | `.github/workflows` and CI configuration |
| `docs` | `README.md`, `CHANGELOG.md`, `AGENTS.md`, documentation files |
| `scripts` | `eng/scripts` |
| `repo` | repository-wide hygiene that does not fit the scopes above |

Examples:

```text
fix(core): pass item type explicitly in InsertData readers
build(packaging): ship symbols, documentation and license files
ci: pin the SDK exactly so lock files stay stable
docs(agents): add repository guidance for agents
test(tests): skip read-only overwrite check on unix
refactor(io): read part names case-insensitively
```

A few `ci`-style commits have no natural scope; a plain `ci: ...` is acceptable, but
prefer a scope when one fits.

## CI gates

| Gate | What it proves |
|---|---|
| `format` | `dotnet format --verify-no-changes`, including encodings and whitespace |
| `build-test` | build with warnings as errors; test run on Linux |
| `codegen-drift` | regenerating the readers produces no diff |
| `pack` | package builds, contains the expected files and stays inside the size budget |
| `aot-smoke` | NativeAOT publish and run of the smoke application (currently non-blocking) |

`aot-smoke` becomes blocking once the smoke project exists and the gate is verified; see
the workflow file for the current state.

## Tests on different platforms

The suite is expected to pass on Linux and Windows. Some tests are explicitly gated
because the platform does not provide the same primitives:

- font metrics for fonts that are not installed and have no embedded metric-compatible
  counterpart;
- read-only file overwrite semantics (`UnauthorizedAccessException` exists on Windows,
  but not for this code path on Unix).

When adding a platform gate, state the reason in a comment next to the attribute.
