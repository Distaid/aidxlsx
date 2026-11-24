# Conditional formatting

Conditional formatting changes how a cell looks when its value satisfies a rule. The rule
itself holds no value: it compares cells at render time, so highlighting stays correct after
the workbook is recalculated. Rules are added to a range or cell and stored per worksheet.

```csharp
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");

sheet.Cell("A1").Value = "Amount";
sheet.Range("A2:A20").Value = 120;

// A rule highlights cells whose value is above 100.
sheet.Range("A2:A20").AddConditionalFormat()
    .WhenGreaterThan(100)
    .Fill.SetBackgroundColor(XlsxColor.FromHtml("#FFC7CE"));

workbook.SaveAs("sales.xlsx");
```

## Adding rules

`AddConditionalFormat()` exists on `IXlsxCell` and on `IXlsxRangeBase`, so cells, ranges, rows,
and columns can all receive a rule. The method creates the rule and returns
`IXlsxConditionalFormat`; a `When*`, `DataBar`, `ColorScale`, or `IconSet` call configures it
and returns the style (`IXlsxStyle`) of the rule, or a builder for the visual rule kinds.

```csharp
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Orders");

sheet.Range("B2:B50").Value = 42;

var cf = sheet.Cell("B2").AddConditionalFormat();
cf.WhenLessThan(0).Font.SetFontColor(XlsxColor.Red);

// One rule can target several ranges; ranges must belong to the rule's worksheet.
var flagged = sheet.Range("C2:C50").AddConditionalFormat();
flagged.WhenIsError().Fill.SetBackgroundColor(XlsxColor.Orange);
flagged.Ranges = [sheet.Range("C2:C50"), sheet.Range("E2:E50")];

// The single-range view is the first area of the rule.
sheet.Cell("D1").Value = flagged.Range.RangeAddress.ToStringRelative(false);

workbook.SaveAs("orders.xlsx");
```

- Assigning `Ranges` replaces all target areas of the rule. The sequence must not be empty and
  every range must be on the same worksheet, otherwise `ArgumentException` is thrown.
  `Range` returns the first area and assigning it collapses the rule back to one area.
- `IXlsxConditionalFormats.Add(rule)` places a rule object in the sheet's collection. Rules
  created with `AddConditionalFormat()` are registered automatically, so the method only matters
  for a rule object obtained from somewhere else, such as a rule that was copied between sheets.
- Rules are not enumerable through the public API. `IXlsxConditionalFormats` only offers `Add`,
  `Remove(predicate)`, and `RemoveAll`, so store the returned object if you need to change the
  rule later.

## Rule kinds

| Goal | API on `IXlsxConditionalFormat` |
|---|---|
| Equality | `WhenEquals`, `WhenNotEquals` |
| Ordering | `WhenGreaterThan`, `WhenLessThan` |
| Inclusive bounds | `WhenEqualOrGreaterThan`, `WhenEqualOrLessThan` |
| Range | `WhenBetween(min, max)`, `WhenNotBetween(min, max)` |
| Text | `WhenContains`, `WhenNotContains`, `WhenStartsWith`, `WhenEndsWith` |
| Blank and error | `WhenIsBlank`, `WhenNotBlank`, `WhenIsError`, `WhenNotError` |
| Duplicates | `WhenIsDuplicate`, `WhenIsUnique` |
| Top and bottom | `WhenIsTop(count, XlsxTopBottomType)`, `WhenIsBottom(count, ...)` |
| Date periods | `WhenDateIs(XlsxTimePeriod)` |
| Formula | `WhenIsTrue(formula)` |
| Visual rules | `DataBar(...)`, `ColorScale()`, `IconSet(...)` |

Every comparison method has a `double` and a `string` overload, and `WhenBetween` /
`WhenNotBetween` take both bounds in the same form.

```csharp
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Rules");

sheet.Range("A2:A20").AddConditionalFormat()
    .WhenBetween(10, 20).Fill.SetBackgroundColor(XlsxColor.LightGreen);

sheet.Range("B2:B20").AddConditionalFormat()
    .WhenContains("urgent").Font.SetBold();

sheet.Range("C2:C20").AddConditionalFormat()
    .WhenIsTop(5, XlsxTopBottomType.Percent).Fill.SetBackgroundColor(XlsxColor.LightBlue);

sheet.Range("D2:D20").AddConditionalFormat()
    .WhenDateIs(XlsxTimePeriod.Today).Font.SetItalic();

sheet.Range("E2:E20").AddConditionalFormat()
    .WhenIsTrue("$A2>100").Fill.SetBackgroundColor(XlsxColor.LightGray);

workbook.SaveAs("rules.xlsx");
```

- `WhenIsTrue` prefixes the formula with `=` when it is missing. The formula is written
  relative to the first cell of the rule's range, so a rule over `E2:E20` uses row `2` in the
  reference and lets Excel shift it down.
- `XlsxTopBottomType` has `Items` (a count of values) and `Percent`; `WhenIsTop(count,
  XlsxTopBottomType.Items)` is the default.
- `XlsxTimePeriod` supports `Yesterday`, `Today`, `Tomorrow`, `InTheLast7Days`, `LastWeek`,
  `ThisWeek`, `NextWeek`, `LastMonth`, `ThisMonth`, and `NextMonth`.
- Above-average and below-average rules are not authorable: there is no `When*` method for
  them, and saving a workbook that contains such a loaded rule throws
  `NotImplementedException`.
- Text rules are case-insensitive matches implemented with `SEARCH`, `LEFT`, and `RIGHT`
  formulas, so wildcard characters in the search text are treated literally.

## Style and differential format

The style of a rule is a differential format (`dxf`): it stores only the properties the rule
changes, in contrast to a cell format which stores a complete style. Font, fill, border,
alignment, number format, and protection can all be overridden.

```csharp
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Risk");

sheet.Range("A2:A20").Value = 5;

var low = sheet.Range("A2:A20").AddConditionalFormat();
low.WhenLessThan(0);

IXlsxStyle dxf = low.Style;
dxf.Font.FontColor = XlsxColor.FromHtml("#9C0006");
dxf.Fill.SetBackgroundColor(XlsxColor.FromHtml("#FFC7CE"));
dxf.Border.SetOutsideBorder(XlsxBorderStyleValues.Thin);
dxf.NumberFormat.Format = "#,##0.00";

workbook.SaveAs("risk.xlsx");
```

- The style is written as a `dxfId` reference for comparison, text, blank, duplicate, and time
  period rules. Data bars, color scales, and icon sets never carry a `dxfId`.
- `IXlsxStyle.IncludeQuotePrefix` throws `NotSupportedException` on a rule style, and
  `IXlsxStyle.Equals` on one throws as well, so do not compare rule styles for equality.
- Assigning a cell style to a rule (`rule.Style = cell.Style`) throws
  `NotSupportedException("Can only copy from dxf.")`. Build the differential format from
  scratch, or read an existing one and modify it.
- A rule without any style properties still matches cells; it simply changes nothing visible.

## Data bars

A data bar draws a bar inside the cell proportional to the value. Its minimum and maximum
define the scale and can be automatic, a number, a percent, a percentile, or a formula.

| `XlsxCFContentType` | Meaning |
|---|---|
| `Minimum`, `Maximum` | Lowest and highest value in the range |
| `Number` | Explicit value |
| `Percent` | Percentage of the range |
| `Percentile` | Percentile of the range |
| `Formula` | Formula result |

```csharp
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Bars");

// Automatic bounds: the lowest and highest value of the range.
sheet.Range("B2:B20").AddConditionalFormat()
    .DataBar(XlsxColor.Green)
    .LowestValue()
    .HighestValue();

// Explicit bounds, a negative-value color, and the bar only (no number in the cell).
sheet.Range("C2:C20").AddConditionalFormat()
    .DataBar(XlsxColor.Green, XlsxColor.Red, showBarOnly: true)
    .Minimum(XlsxCFContentType.Number, 0)
    .Maximum(XlsxCFContentType.Number, 500);

workbook.SaveAs("bars.xlsx");
```

- `LowestValue()` sets the minimum type to `Minimum` and `HighestValue()` sets the maximum
  type to `Maximum`. Use `Minimum(type, value)` and `Maximum(type, value)` for explicit values;
  each also has a `string` overload for formulas.
- `showBarOnly` hides the cell value and shows the bar only. `ShowBarOnly` is read-only on the
  rule object, so pass the flag to `DataBar`.
- The negative-value color is stored in the `x14` extension. Gradient rendering is always
  written as enabled, and bar direction is not exposed, so both follow Excel defaults.
- Bar lengths are computed by the spreadsheet application from the values and thresholds;
  Aid.Xlsx stores no bar geometry.

## Color scales

A color scale shades cells between two or three colors according to their value.

```csharp
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Scale");

// Two colors: lowest value in white, highest in red.
sheet.Range("A2:A20").AddConditionalFormat()
    .ColorScale()
    .LowestValue(XlsxColor.White)
    .HighestValue(XlsxColor.Red);

// Three colors with an explicit midpoint.
sheet.Range("B2:B20").AddConditionalFormat()
    .ColorScale()
    .Minimum(XlsxCFContentType.Percent, 0, XlsxColor.LightYellow)
    .Midpoint(XlsxCFContentType.Percentile, 50, XlsxColor.Yellow)
    .HighestValue(XlsxColor.Red);

workbook.SaveAs("scale.xlsx");
```

- `ColorScale()` returns the minimum-threshold builder; `LowestValue(color)` and
  `Minimum(type, value, color)` both return the next builder, and `HighestValue(color)` or
  `Maximum(type, value, color)` end the chain.
- A three-color scale needs exactly three thresholds and three colors. Calling `Midpoint` adds
  the middle one; the resulting `IXlsxCFColorScaleMax` finishes the scale.
- Colors accept the same values as cell colors: named colors, `XlsxColor.FromArgb`,
  `XlsxColor.FromHtml`, indexed, and theme colors.

## Icon sets

An icon set places one of three, four, or five icons in the cell according to threshold
values.

| Icons | `XlsxIconSetStyle` members |
|---|---|
| 3 | `ThreeArrows`, `ThreeArrowsGray`, `ThreeFlags` |
| 3 | `ThreeTrafficLights1`, `ThreeTrafficLights2`, `ThreeSigns` |
| 3 | `ThreeSymbols`, `ThreeSymbols2` |
| 4 | `FourArrows`, `FourArrowsGray`, `FourRedToBlack` |
| 4 | `FourRating`, `FourTrafficLights` |
| 5 | `FiveArrows`, `FiveArrowsGray`, `FiveRating`, `FiveQuarters` |

```csharp
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Icons");

sheet.Range("A2:A20").AddConditionalFormat()
    .IconSet(XlsxIconSetStyle.ThreeArrows)
    .AddValue(XlsxCFIconSetOperator.EqualOrGreaterThan, 0, XlsxCFContentType.Percent)
    .AddValue(XlsxCFIconSetOperator.EqualOrGreaterThan, 33, XlsxCFContentType.Percent)
    .AddValue(XlsxCFIconSetOperator.GreaterThan, 67, XlsxCFContentType.Percent);

sheet.Range("B2:B20").AddConditionalFormat()
    .IconSet(XlsxIconSetStyle.FiveRating, reverseIconOrder: true, showIconOnly: true)
    .AddValue(XlsxCFIconSetOperator.EqualOrGreaterThan, 0, XlsxCFContentType.Percent)
    .AddValue(XlsxCFIconSetOperator.EqualOrGreaterThan, 20, XlsxCFContentType.Percent)
    .AddValue(XlsxCFIconSetOperator.EqualOrGreaterThan, 40, XlsxCFContentType.Percent)
    .AddValue(XlsxCFIconSetOperator.EqualOrGreaterThan, 60, XlsxCFContentType.Percent)
    .AddValue(XlsxCFIconSetOperator.EqualOrGreaterThan, 80, XlsxCFContentType.Percent);

workbook.SaveAs("icons.xlsx");
```

- Add one threshold per icon, in ascending order. The first threshold is usually `0` percent,
  and the last one uses `GreaterThan` because the top icon covers everything above it.
- `XlsxCFIconSetOperator` has `GreaterThan` and `EqualOrGreaterThan`, written as the `gte`
  flag of each threshold. Threshold types come from `XlsxCFContentType`.
- `reverseIconOrder` flips the icon order and `showIconOnly` hides the value. `ReverseIconOrder`
  and `ShowIconOnly` are read-only on the rule object, so pass both flags to `IconSet`.
- The icon style must be set on the `IconSet` call. Calling `IconSet` again clears previously
  added thresholds.

## Priority, ordering, and removal

Rules have no public priority property. Their precedence follows the order in which they were
added, and Aid.Xlsx renumbers the `priority` attribute from 1 when the workbook is saved.

```csharp
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Priority");

var high = sheet.Range("A2:A20").AddConditionalFormat();
high.WhenGreaterThan(100).Fill.SetBackgroundColor(XlsxColor.Red);
high.SetStopIfTrue();

// Second rule for the same cells: used only where the first rule did not match.
sheet.Range("A2:A20").AddConditionalFormat()
    .WhenGreaterThan(50).Fill.SetBackgroundColor(XlsxColor.Yellow);

// Remove selected rules, or all of them.
sheet.ConditionalFormats.Remove(cf => cf.ConditionalFormatType == XlsxConditionalFormatType.Top10);
sheet.ConditionalFormats.RemoveAll();

// Clear only the rules that intersect a range.
sheet.Range("A2:A20").Clear(XlsxClearOptions.ConditionalFormats);

workbook.SaveAs("priority.xlsx");
```

- When several rules target the same range, the first matching rule wins. `SetStopIfTrue()`
  marks a rule so Excel stops evaluating later rules as soon as it matches, and
  `SetStopIfTrue(false)` clears the flag. The flag is stored per rule and only matters when a
  later rule overlaps the same cells.
- Rules with an identical target range are written under one `conditionalFormatting` element,
  in list order. Rules with different ranges keep their relative order and their own
  `priority` values.
- `Remove(predicate)` filters by any public rule property, for example
  `ConditionalFormatType`, `Operator`, `StopIfTrue`, or `Range`. `RemoveAll()` clears the sheet.
- Clearing a range removes or shrinks overlapping rules: a rule covering `A1:C10` that is
  cleared for `B1:B10` keeps `A1:A10` and `C1:C10` as two areas, and disappears when no area is
  left. Cell and range clearing with `XlsxClearOptions.ConditionalFormats` follows the same
  rule.
- `SaveOptions.ConsolidateConditionalFormatRanges` defaults to `true` and merges adjacent
  identical rules that do not intersect any other rule, so the number of rules in the saved
  file can be smaller than the number added at runtime.

## Interaction with other features

Conditional formatting is stored as worksheet data, so table cells and pivot-adjacent ranges
behave like ordinary cells: "banded rows" in a table style are table formatting rather than
conditional formatting, and no API creates table conditional rules.

- Pivot table conditional rules are loaded into the pivot table part and are not exposed
  through `IXlsxWorksheet.ConditionalFormats`; the public API cannot read or author them.
- `IXlsxWorksheet.CopyTo` copies conditional rules together with the cells, preserving relative
  formulas. `IXlsxConditionalFormat.CopyTo(worksheet)` copies a single rule to another
  worksheet and throws `InvalidOperationException` when the target is the source sheet.
- Copying a cell to another position copies the overlapping rule; copying a single cell inside
  an existing rule's area merges the copied cell into that rule instead of duplicating it.
- Deleting or inserting rows and columns shifts rule areas and relative formulas, and a rule is
  dropped when its last area is deleted.
- The file stores only what Excel needs to evaluate the rule. Aid.Xlsx writes the rule kind,
  operator, thresholds, colors, and formulas but no rendered colors, so a viewer that ignores
  conditional formatting shows the plain cell style.

Rules are part of worksheet content, so a rule survives a round trip even though the public
API cannot list the rules of a sheet. For the cell style a rule overrides, see
[Formatting and worksheet features](formatting-and-worksheets.md); for table and pivot
interaction, see [Tables and data import](tables-and-data-import.md) and
[Pivot tables](pivot-tables.md).
