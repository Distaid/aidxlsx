# Aid.Xlsx.Generators

This project contains the Roslyn incremental generators for Aid.Xlsx. It targets `netstandard2.0`
for the compiler host. The main NuGet package installs its assembly as a C# analyzer.

| Directory | Responsibility |
| --- | --- |
| [TableMapping](TableMapping/) | Discovers models annotated with `XlsxTable` and generates typed row mappers for `InsertMappedData` and `InsertMappedTable` |
| [WorkbookMetadata](WorkbookMetadata/) | Generates enum description and built-in table theme lookups while compiling the workbook assembly |
| [Diagnostics](Diagnostics/) | Tracks diagnostic IDs and analyzer releases |

The NativeAOT smoke project exercises a generated row mapper. Package CI also builds the smoke
project using the analyzer from the packed NuGet. [Smoke project](../../tests/Aid.Xlsx.AotSmoke/)
