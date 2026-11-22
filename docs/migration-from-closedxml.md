# Migrate a workbook application to Aid.Xlsx

Reference the `Aid.Xlsx` package. Switch `XL*` to `Xlsx*` and `IXL*` to `IXlsx*`
in type names, then import the namespace for each type's domain. The general exception
type is `Aid.Xlsx.Exceptions.XlsxException`.

| Type | Namespace |
| --- | --- |
| `XlsxWorkbook`, `SaveOptions` | `Aid.Xlsx.Workbooks` |
| `IXlsxWorksheet` | `Aid.Xlsx.Worksheets` |
| `IXlsxCell`, `XlsxCellValue` | `Aid.Xlsx.Cells` |
| `IXlsxRange` | `Aid.Xlsx.Ranges` |
| `IXlsxTable`, `XlsxTableTheme` | `Aid.Xlsx.Tables` |
| `IXlsxStyle` | `Aid.Xlsx.Styles` |
| `XlsxColor` | `Aid.Xlsx.Styles.Colors` |
| `XlsxColumnAttribute`, `XlsxTableAttribute` | `Aid.Xlsx.Attributes` |

```csharp
using Aid.Xlsx.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
workbook.Worksheet(1).Cell("B2").Value = 42;
workbook.SaveAs("output.xlsx");
```

Recompile consumers after updating type names. Assemblies compiled against ClosedXML
keep their original assembly references. Aid.Xlsx exports the new names without aliases.

For annotated row models in trimmed applications, use `InsertMappedTable` or
`InsertMappedData`. See the [NativeAOT guide](native-aot.md).
