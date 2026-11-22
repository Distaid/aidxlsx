# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- Use checked-in compatibility baselines and isolate the optional comparison benchmark from the main solution.
- Generate built-in table theme and protection algorithm lookups at compile time.
- Resolve and write workbook parts and relationships through the OPC container.
- Remove the runtime Open XML SDK dependency. `SaveOptions.ValidatePackage` checks
  workbook relationships and supported XML parts. The tests validate output against the SDK schemas.
- Rename `XL*` and `IXL*` types to `Xlsx*` and `IXlsx*`. Public types now live in
  directory-based namespaces such as `Aid.Xlsx.Model.Workbooks`, `Aid.Xlsx.Model.Worksheets`,
  and `Aid.Xlsx.Model.Grid.Cells`.
  Existing consumers must update their imports and type names.
- Write formula calculation chains, styles, shared strings, comments and pivot cache records through XML streams.
- Read and write core, custom and extended workbook properties directly through OPC, retaining metadata outside the workbook model.
- Read theme colors and generate default themes from XML. The streaming worksheet data reader uses its own cell and formula types.
- Stream worksheet cells directly through XML readers and writers without SDK reflection.
- Read and write worksheet layout, print settings, hyperlinks, validations, and sparklines directly through XML.
- Read and write table definitions and auto-filters directly through XML.
- Read and write pivot table and cache definitions directly through XML.
- Read and write worksheet conditional formatting rules and data bar settings directly through XML.
- Read comment text with the same XML parser as rich text in cells.
- Write legacy comment shapes directly through XML while preserving the drawing's other shapes.
- Read and write picture anchors directly through DrawingML and preserve unrelated drawing elements when deleting pictures.
- Read and write workbook settings, sheets, views, and defined names directly through SpreadsheetML.
- Open and save workbook ZIP containers with the built-in OPC layer, preserving untouched parts and their relationships.
- Resolve workbook and worksheet parts directly through OPC relationships when loading Strict and Transitional files.
- Build the formula parser from repository sources and ship it inside `Aid.Xlsx`.
- Organize the five source projects by responsibility and document each in its own README.
  The workbook model, SpreadsheetML codecs, calculation engine, and BIFF8 reader have separate directories.

### Added

- Include a CycloneDX dependency SBOM and SHA-256 checksums in the release artifacts.
- Convert Excel 97–2003 BIFF8 worksheet data to `.xlsx` from paths or streams without Office automation.
- Convert an `.xls` file to `.xlsx` in the same directory by passing only the source path.
- Convert supported BIFF8 formulas to XLSX while retaining their cached results.
- Preserve BIFF8 formula references to individual sheets in the same workbook.
- Preserve BIFF8 cell fonts, alignment, borders, fills, protection flags, and palette colors during conversion.
- Generate reflection-free row mappers for `InsertMappedTable` and `InsertMappedData` from `[XlsxTable]` models.
- Round-trip a workbook in fully trimmed and NativeAOT smoke applications.
- Add an OPC ZIP container with content types, internal and external relationships, and byte-preserving copies of unchanged parts.
- Initial repository files: `.gitignore`, the solution file `Aid.Xlsx.slnx` with the
  imported projects and the project documentation (`README.md`, this changelog).

### Fixed

- Preserve an existing pivot cache's refresh-on-open setting when saving a workbook.
- Install `Aid.Xlsx` without resolving a separate `Aid.Xlsx.Opc` package.
- `InsertData`/`InsertTable` pass the row type to the readers explicitly instead of
  inferring it from the materialized row collection, which returned `object` and
  produced empty ranges. Rows of different runtime types are no longer mixed in one
  grouped snapshot.

### Planned

- `0.1.0-preview1` — the first public preview on NuGet. Publication and the release date are pending.

[Unreleased]: https://github.com/Distaid/aidxlsx/commits/main
