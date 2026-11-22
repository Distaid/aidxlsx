# Aid.Xlsx.Tests

The NUnit suite tests the workbook API and the libraries it uses to process files.
Namespaces follow the directory path under `Aid.Xlsx.Tests`.

| Directory | Coverage |
| --- | --- |
| `Contracts/` | Public API naming and namespace conventions |
| `Model/` | Workbook, worksheet, grid, and feature behavior |
| `Calculation/` | Formula evaluation, functions, references, and recalculation |
| `DataImport/` | Row readers and source-generated table mapping |
| `Conversion/Biff8/` | Excel 97–2003 workbook conversion |
| `SpreadsheetML/` | XLSX/XLSM part readers, writers, and workbook round-trips |
| `Opc/` | Package container, relationships, and XML readers |
| `Infrastructure/` and `Rendering/` | Library utilities, fonts, and image metadata |
| `TestSupport/` | NUnit helpers, resource loading, and workbook comparisons |

`Resource/` holds the embedded workbook, image, font, and text fixtures.
Workbook comparisons use these fixtures as references and use the Open XML SDK
for schema validation. The BIFF8 fixtures are in `Resource/Biff8/` and are also
used by the NativeAOT smoke project.

From the repository root, run the suite or filter it by namespace.

```powershell
dotnet test tests/Aid.Xlsx.Tests -c Release
dotnet test tests/Aid.Xlsx.Tests -c Release --filter "FullyQualifiedName~Aid.Xlsx.Tests.SpreadsheetML"
```
