# SpreadsheetML schema for code generation

The reader generator uses `sml.xsd` from ECMA-376, Part 4, fifth edition (December 2016).
The schema is inside `OfficeOpenXML-XMLSchema-Transitional.zip` in the official download.
The archive carries the ECMA copyright and distribution terms. [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)

Run `dotnet run --project src/Aid.Xlsx.CodeGen -- all` from the repository root. The tool
downloads the official archive on first use, verifies the extracted schema against
`SHA256SUMS`, and caches it in `artifacts/codegen/sml.xsd` for offline regeneration.
