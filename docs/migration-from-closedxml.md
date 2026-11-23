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

## Known gaps and replacements

These mappings apply to both Aid.Xlsx 1.x (.NET 10) and 2.x (.NET 11).

| ClosedXML member | Aid.Xlsx equivalent or migration path |
| --- | --- |
| `XLWorkbook.DefaultStyle` | Set the default on the workbook instance with `workbook.Style`. Cells inherit it wherever no closer format overrides it. In 1.0.0, set `worksheet.Style` on each new worksheet to work around an inheritance bug. |
| `XLWorkbook.GetAnchorFromImageId`, `GetAnchorFromImageIndex`, `GetImageRelIdFromAnchor`, `GetPropertiesFromAnchor` | Use `worksheet.Pictures` and the `IXlsxPicture` properties for IDs, placement, cell anchors, offsets, and dimensions of supported pictures. Code that needs raw DrawingML anchors or relationship IDs can read the package separately. |
| `IXLSparkline.IsValid` | After a structural edit, retrieve the sparkline again from the current cell via `cell.Sparkline`. Test `sparkline?.SourceData is not null` if you need a source range that the current workbook can resolve. |
| `IXLCell.SetDataValidation()` | Use `cell.GetDataValidation()` to reuse or create a rule. Check `cell.HasDataValidation` before calling it when you only want to read an existing rule. Use `cell.CreateDataValidation()` to replace a rule. |
| `IXLRangeBase.SetDataValidation()` | Use `range.GetDataValidation()` to look up an existing rule or `range.CreateDataValidation()` for a new rule. `GetDataValidation()` on a range can return `null`. If your code relied on copying settings from an overlapping rule, capture and copy those settings explicitly. |
| `XLColor.FromHtml(string)` | Use `XlsxColor.FromHtml("#336699")` for RGB hex. `FromHexRgb` is an internal XML helper. |

Use `workbook.Style` to set a baseline font and alignment when creating a workbook. Both
`CellsUsed(predicate)` and `CellsUsed(options, predicate)` are inherited by `IXlsxWorksheet`
from `IXlsxRangeBase`. They apply the predicate during used-cell enumeration.

```csharp
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
workbook.Style.Font.FontName = "Arial";
workbook.Style.Alignment.Horizontal = XlsxAlignmentHorizontalValues.Center;

var worksheet = workbook.AddWorksheet("Export");
worksheet.Cell("A1").Value = "Invoice";

foreach (var cell in worksheet.CellsUsed(c => c.DataType == XlsxDataType.Text))
    Console.WriteLine(cell.GetText());
```

Use `XlsxCellsUsedOptions.AllContents` for contents, types, and comments. Combine it with
`XlsxCellsUsedOptions.NormalFormats` to also visit cells whose format differs from what they
inherit. The [compatibility guide](compatibility.md#template-export-regression-tests) lists
the regression tests for editing and saving templates.

For annotated row models in trimmed applications, use `InsertMappedTable` or
`InsertMappedData`. See the [NativeAOT guide](native-aot.md).
