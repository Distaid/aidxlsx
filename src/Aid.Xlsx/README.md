# Aid.Xlsx

This project provides the workbook API, calculation engine, and format conversion. Start with
[`XlsxWorkbook`](Model/Workbooks/XlsxWorkbook.cs) to work with an XLSX or XLSM file and
[`XlsConverter`](Conversion/XlsConverter.cs) to import an Excel 97–2003 workbook.

| Directory | Responsibility |
| --- | --- |
| [Model](Model/) | Workbook and worksheet state, grid coordinates and cells, styles, tables, pivots, drawings, and other workbook features |
| [SpreadsheetML](SpreadsheetML/) | Resolves workbook parts through OPC and reads or writes their XML. Each part family has its own directory. |
| [Calculation](Calculation/) | Formula evaluation, dependency tracking, and worksheet functions |
| [Conversion](Conversion/) | BIFF8 records and formulas, plus the compound-file reader that opens legacy workbooks |
| [DataImport](DataImport/) | Row readers, table mapping attributes, and registration of generated mappers |
| [Rendering](Rendering/) | Font metrics, embedded fonts, and image metadata readers |
| [Infrastructure](Infrastructure/) | Shared collections, caches, and extensions |

Loading and saving start in `Model/Workbooks/`. The SpreadsheetML codecs populate or serialize
the model. The OPC project handles ZIP storage and relationships, and the formula project
supplies parsed formulas to the calculation engine. Namespaces match the source directories,
including `Aid.Xlsx.Model.Grid.Cells`, `Aid.Xlsx.Model.Styles`, and `Aid.Xlsx.Model.Workbooks`.

Run `dotnet build src/Aid.Xlsx -c Release` from the repository root. The repository README covers
package usage and directs contributors to the test and NativeAOT guides. [Repository README](../../README.md)
