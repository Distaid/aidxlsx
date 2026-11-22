# Migrate a workbook application to Aid.Xlsx

Reference the `Aid.Xlsx` package. Switch `XL*` to `Xlsx*` and `IXL*` to `IXlsx*`
in type names, then import the namespace for each type's domain. The general exception
type is `Aid.Xlsx.Model.Exceptions.XlsxException`.

| Type | Namespace |
| --- | --- |
| `XlsxWorkbook`, `SaveOptions` | `Aid.Xlsx.Model.Workbooks` |
| `IXlsxWorksheet` | `Aid.Xlsx.Model.Worksheets` |
| `IXlsxCell`, `XlsxCellValue` | `Aid.Xlsx.Model.Grid.Cells` |
| `IXlsxRange` | `Aid.Xlsx.Model.Grid.Ranges` |
| `IXlsxTable`, `XlsxTableTheme` | `Aid.Xlsx.Model.Tables` |
| `IXlsxStyle` | `Aid.Xlsx.Model.Styles` |
| `XlsxColor` | `Aid.Xlsx.Model.Styles.Colors` |
| `XlsxColumnAttribute`, `XlsxTableAttribute` | `Aid.Xlsx.DataImport.Attributes` |

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("input.xlsx");
workbook.Worksheet(1).Cell("B2").Value = 42;
workbook.SaveAs("output.xlsx");
```

Recompile consumers after updating type names. Assemblies compiled against ClosedXML
keep their original assembly references. Aid.Xlsx exports the new names without aliases.

For annotated row models in trimmed applications, use `InsertMappedTable` or
`InsertMappedData`. See the [NativeAOT guide](native-aot.md).
