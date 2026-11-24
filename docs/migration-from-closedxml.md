# Migrate a workbook application to Aid.Xlsx

Aid.Xlsx follows the ClosedXML 0.106 API shape but uses its own OPC and
SpreadsheetML layer, without the Open XML SDK runtime dependency. The workbook model is
recognizably the same, so migration is mostly renames and namespace imports.

## What changes for consumers

- Reference the `Aid.Xlsx` package instead of `ClosedXML`.
- Recompile against Aid.Xlsx. The assembly name changes, and types are not forwarded:
  a binary compiled against ClosedXML keeps its ClosedXML reference and does not work
  with Aid.Xlsx.
- Rename `XL*` types to `Xlsx*` and `IXL*` types to `IXlsx*`, then import the namespace
  that owns each domain.
- Add `using Aid.Xlsx.Model.Workbooks;` for `XlsxWorkbook`. A file that used a single
  ClosedXML namespace usually needs two or three Aid.Xlsx namespaces.

## Namespace map

| Type | Namespace |
| --- | --- |
| `XlsxWorkbook`, `LoadOptions`, `SaveOptions`, `XlsxCalculateMode`, `XlsxReferenceStyle` | `Aid.Xlsx.Model.Workbooks` |
| `IXlsxWorksheet`, `IXlsxWorksheets`, `IXlsxSheetView`, `IXlsxOutline`, `XlsxWorksheetVisibility` | `Aid.Xlsx.Model.Worksheets` |
| `IXlsxCell`, `XlsxCellValue`, `XlsxDataType`, `XlsxCellsUsedOptions`, `XlsxCellSetValueBehavior` | `Aid.Xlsx.Model.Grid.Cells` |
| `IXlsxRange`, `IXlsxRangeBase`, `XlsxScope`, `XlsxShiftDeletedCells` | `Aid.Xlsx.Model.Grid.Ranges` |
| `XlsxSortOrder` and sort helpers | `Aid.Xlsx.Model.Grid.Ranges.Sort` |
| `IXlsxRow`, `IXlsxColumn` | `Aid.Xlsx.Model.Grid.Rows`, `Aid.Xlsx.Model.Grid.Columns` |
| `IXlsxAddress`, `XlsxReferenceStyle` | `Aid.Xlsx.Model.Grid.Coordinates` |
| `IXlsxStyle`, `XlsxAlignmentHorizontalValues` | `Aid.Xlsx.Model.Styles` |
| `XlsxColor`, `XlsxThemeColor` | `Aid.Xlsx.Model.Styles.Colors` |
| `IXlsxTable`, `XlsxTableTheme` | `Aid.Xlsx.Model.Tables` |
| `IXlsxDataValidation`, `XlsxOperator`, `XlsxAllowedValues`, `XlsxErrorStyle`, and the criteria types (`XlsxWholeNumberCriteria`, `XlsxDecimalCriteria`, `XlsxDateCriteria`, `XlsxTimeCriteria`, `XlsxTextLengthCriteria`) | `Aid.Xlsx.Model.DataValidation` |
| `IXlsxConditionalFormat`, data bar and icon set types | `Aid.Xlsx.Model.ConditionalFormats` |
| `IXlsxComment` | `Aid.Xlsx.Model.Comments` |
| `IXlsxRichText`, rich text runs | `Aid.Xlsx.Model.RichText` |
| `XlsxHyperlink` | `Aid.Xlsx.Model.Hyperlinks` |
| `IXlsxPicture`, `XlsxPicturePlacement`, `XlsxPictureFormat` | `Aid.Xlsx.Model.Drawings` |
| `IXlsxSparkline`, `IXlsxSparklineGroup`, `XlsxSparklineType` | `Aid.Xlsx.Model.Sparklines` |
| `IXlsxPivotTable`, `IXlsxPivotCache` | `Aid.Xlsx.Model.PivotTables` |
| `IXlsxDefinedName`, `IXlsxDefinedNames` | `Aid.Xlsx.Model.DefinedNames` |
| `IXlsxPageSetup`, header and footer types | `Aid.Xlsx.Model.PageSetup` |
| `IXlsxSheetProtection`, `XlsxSheetProtectionElements` | `Aid.Xlsx.Model.Protection` |
| `XlsxTableAttribute`, `XlsxColumnAttribute`, `XlsxIgnoreAttribute` | `Aid.Xlsx.DataImport.Attributes` |
| `InsertMappedData`, `InsertMappedTable` | `Aid.Xlsx.DataImport.Mapping` |
| `XlsConverter` | `Aid.Xlsx.Conversion` |
| `XlsxException`, `EmptyTableException`, `InvalidReferenceException` | `Aid.Xlsx.Model.Exceptions` |

## Rename examples

| ClosedXML | Aid.Xlsx |
| --- | --- |
| `XLWorkbook` | `XlsxWorkbook` |
| `IXLWorksheet` | `IXlsxWorksheet` |
| `IXLCell` | `IXlsxCell` |
| `IXLRange` | `IXlsxRange` |
| `IXLRangeBase` | `IXlsxRangeBase` |
| `IXLTable` | `IXlsxTable` |
| `IXLStyle` | `IXlsxStyle` |
| `XLColor` | `XlsxColor` |
| `IXLDataValidation` | `IXlsxDataValidation` |
| `IXLConditionalFormat` | `IXlsxConditionalFormat` |
| `IXLComment` | `IXlsxComment` |
| `IXLRichText` | `IXlsxRichText` |
| `IXLPicture` | `IXlsxPicture` |
| `IXLSparklineGroup` | `IXlsxSparklineGroup` |
| `IXLPivotTable` | `IXlsxPivotTable` |
| `IXLDefinedName` | `IXlsxDefinedName` |
| `XLEmptyTableException` | `EmptyTableException` |

## API mapping

| Task | ClosedXML | Aid.Xlsx |
| --- | --- | --- |
| Create a workbook | `new XLWorkbook()` | `new XlsxWorkbook()` |
| Open a file | `new XLWorkbook(path)` | `new XlsxWorkbook(path)` |
| Open with options | constructor overload | `new XlsxWorkbook(path, new LoadOptions())` |
| Save to the source | `workbook.Save()` | `workbook.Save()` |
| Save as | `workbook.SaveAs(path)` | `workbook.SaveAs(path)` |
| Add a worksheet | `workbook.Worksheets.Add("Name")` | `workbook.AddWorksheet("Name")` or `workbook.Worksheets.Add("Name")` |
| Look up a worksheet | `workbook.Worksheet("Name")` | `workbook.Worksheet("Name")`, `TryGetWorksheet(...)` |
| Delete a worksheet | `worksheet.Delete()` | `workbook.Worksheets.Delete("Name")` |
| Copy a worksheet | `worksheet.CopyTo(workbook, "Name")` | `worksheet.CopyTo(workbook, "Name")` |
| Cell by address | `sheet.Cell("A1")` | `sheet.Cell("A1")` |
| Cell by indexes | `sheet.Cell(1, 1)` | `sheet.Cell(1, 1)` |
| Value | `cell.Value = 42` | `cell.Value = 42` (implicit `XlsxCellValue` conversion) |
| Formula | `cell.FormulaA1 = "SUM(A1:A3)"` | `cell.FormulaA1 = "SUM(A1:A3)"` |
| Cached result | `cell.CachedValue` | `cell.CachedValue` |
| Typed read | `cell.GetValue<int>()`, `GetDouble()` | `cell.GetValue<int>()`, `GetDouble()` |
| Display text | `cell.GetFormattedString()` | `cell.GetFormattedString()` and `GetFormattedString(CultureInfo)` |
| Range | `sheet.Range("A1:C3")` | `sheet.Range("A1:C3")` |
| Used cells | `range.CellsUsed(predicate)` | `range.CellsUsed(predicate)` or `CellsUsed(options, predicate)` |
| Merge | `range.Merge()` | `range.Merge()` |
| Insert rows | `range.InsertRowsAbove(n)` | `range.InsertRowsAbove(n)` |
| Delete rows or columns | `range.Rows().Delete()` | `range.Rows().Delete()`, `range.Columns().Delete()` |
| Sort | `range.Sort()` | `range.Sort()`, `range.Sort(column, XlsxSortOrder)` |
| Style | `cell.Style.Font.Bold = true` | `cell.Style.Font.Bold = true` |
| Number format | `cell.Style.NumberFormat.Format = "0.00"` | same |
| Column width | `sheet.Column(1).Width = 20` | same, plus `AdjustToContents()` |
| Freeze panes | `sheet.SheetView.FreezeRows(1)` | same |
| Table | `range.CreateTable("Name")` | `range.CreateTable("Name")` |
| Find table | `workbook.Table("Name")` | `workbook.Table("Name")` |
| Auto-filter | `range.SetAutoFilter()` | `range.SetAutoFilter()` |
| Conditional format | `range.AddConditionalFormat().WhenLessThan(...)` | same |
| Data validation, new rule | `range.SetDataValidation()` | `range.CreateDataValidation()` |
| Data validation, reuse | `range.SetDataValidation()` | `range.GetDataValidation()`, `HasDataValidation` |
| Comment | `cell.CreateComment()` | `cell.CreateComment()`, `GetComment()`, `HasComment` |
| Rich text | `cell.CreateRichText()` | `cell.CreateRichText()`, `GetRichText()`, `HasRichText` |
| Hyperlink | `cell.SetHyperlink(link)` | `cell.CreateHyperlink()`, `GetHyperlink()`, `HasHyperlink` |
| Picture | `sheet.AddPicture(stream)` | `sheet.AddPicture(stream)`, `Pictures` collection |
| Sparkline | `sheet.SparklineGroups.Add(...)` | `sheet.SparklineGroups.Add(...)` |
| Pivot table | `range.CreatePivotTable(cell, "Name")` | `range.CreatePivotTable(cell, "Name")` |
| Defined name | `workbook.NamedRanges.Add(...)` | `workbook.DefinedNames.Add(...)`, `range.AddToNamed(...)` |
| Sheet protection | `sheet.Protect(password)` | `sheet.Protect(password)`, `Unprotect(password)` |
| Color from hex | `XLColor.FromHtml("#336699")` | `XlsxColor.FromHtml("#336699")` |

## Behavior differences

- **Default style.** `workbook.Style` applies to worksheets added afterwards. In 1.0.0
  this inheritance had a bug; setting `worksheet.Style` explicitly worked around it.
  The fix shipped in 2.0.0 and 1.0.1, so the workaround is no longer needed.
- **Data validation.** `SetDataValidation()` became three operations: `CreateDataValidation()`
  replaces or creates a rule, `GetDataValidation()` looks up an existing rule (may return
  `null`), and `HasDataValidation` checks without creating. Copying settings from an
  overlapping rule is no longer implicit; capture and copy the settings explicitly.
- **Sparklines.** `IXLSparkline.IsValid` has no direct replacement. Read
  `cell.Sparkline` again after a structural edit and test `sparkline?.SourceData` for a
  source range that still resolves.
- **Pictures.** Raw DrawingML anchor helpers from ClosedXML have no counterpart. Use the
  `IXlsxPicture` properties (`TopLeftCell`, `BottomRightCell`, `Left`, `Top`, `Width`,
  `Height`, `Placement`) and `GetOffset`. Raw anchors and relationship IDs require reading
  the package separately.
- **Exceptions.** The general exception type is `XlsxException`. Argument problems throw
  the standard `ArgumentException` family, including `ArgumentNullException` for null
  names. An empty table input throws `EmptyTableException`; an unresolvable reference
  throws `InvalidReferenceException`.
- **Obsolete aliases.** `NamedRange` and `NamedRanges` remain as obsolete aliases of
  `DefinedName` and `DefinedNames`. New code should use the `DefinedName` names.
- **Nullable annotations.** Several members now carry nullable annotations. Existing
  code that assumed non-null values may need null checks or the null-forgiving operator
  after migration.
- **Charts.** Chart creation and editing are not exposed. Charts in existing workbooks
  are preserved as package data.
- **Encrypted workbooks.** Workbooks encrypted with a password to open are not supported.
  Sheet and workbook protection are supported and are not encryption.

## Step-by-step checklist

1. Replace the package reference and restore.
2. Rewrite type names (`XL` → `Xlsx`, `IXL` → `IXlsx`) with search and replace in source,
   not in generated or binary files.
3. Add the namespace imports from the table above; remove `ClosedXML.*` usings.
4. Fix the compile errors in this order: workbook and worksheet creation, cell and range
   access, styles, tables, then the optional features (validation, conditional formatting,
   comments, pictures, pivot tables).
5. Replace `SetDataValidation()` call sites with `CreateDataValidation()` or
   `GetDataValidation()` depending on intent, and add `HasDataValidation` checks where
   the old code only read settings.
6. Review nullable warnings; they usually point to worksheet or table lookups that can
   return `null`.
7. Run the application against a copy of a real workbook, save, and reopen the result.
   Check the features listed in [Compatibility](compatibility.md) that you rely on.
8. If you use NativeAOT or trimming, replace reflection-based imports with
   `InsertMappedData`/`InsertMappedTable`; see [NativeAOT](native-aot.md).

## Before and after

```csharp
// Before: ClosedXML types and namespaces.
using ClosedXML.Excel;

using var workbook = new XLWorkbook();
var sheet = workbook.Worksheets.Add("Report");
sheet.Cell("A1").Value = "Item";
sheet.Cell("B1").Value = "Units";
sheet.Cell("A2").Value = "Pears";
sheet.Cell(2, 2).Value = 12;
sheet.Range("A1:B1").Style.Font.Bold = true;
sheet.Range("B2:B10").SetDataValidation().WholeNumber.Between(1, 100);
workbook.SaveAs("report.xlsx");
```

```csharp
// After: Aid.Xlsx types and namespaces.
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");
sheet.Cell("A1").Value = "Item";
sheet.Cell("B1").Value = "Units";
sheet.Cell("A2").Value = "Pears";
sheet.Cell(2, 2).Value = 12;
sheet.Range("A1:B1").Style.Font.Bold = true;
sheet.Range("B2:B10").CreateDataValidation().WholeNumber.Between(1, 100);
workbook.SaveAs("report.xlsx");
```

## How to verify the migration

- Build with warnings as errors: the repository treats nullable and analyzer warnings as
  errors, and your application should do the same at least once after migration.
- Compare the produced file by reopening it with Aid.Xlsx and with a spreadsheet
  application. The repository's [compatibility guide](compatibility.md) lists the checks
  it runs on its own corpus.
- If you published an API surface of your own, keep the Aid.Xlsx API baseline in mind:
  `tests/Aid.Xlsx.ApiCompat/Baselines/PublicApi.txt` describes every public member used
  in these guides.

## Known gaps

- Package and DrawingML helpers that exposed Open XML SDK types are absent from the
  workbook API.
- Some low-level types differ: style cache types, arithmetic helpers, obsolete validation
  APIs, range collections, sparkline state, and picture disposal.
- Chart creation, macros, and encrypted packages are out of scope. Macros are preserved
  in `.xlsm` files as opaque data.
- The BIFF8 converter is an addition, not a ClosedXML equivalent. See
  [Workbook files and streams](workbook-files.md#convert-a-legacy-xls-workbook).

## See also

- [Compatibility](compatibility.md)
- [Getting started](getting-started.md)
- [Troubleshooting](troubleshooting.md)
