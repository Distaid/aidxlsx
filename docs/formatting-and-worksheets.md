# Formatting and worksheet features

A cell's appearance comes from its style: font, fill, border, alignment, number format, and
protection flags. The worksheet adds layout (rows, columns, panes) and print settings on top.
Cell values and addresses are covered in [Cells and ranges](cells-and-ranges.md), sheet
protection in [Protection](protection.md). Rules that react to values live in
[Conditional formatting](conditional-formatting.md), rules that restrict input in
[Data validation](data-validation.md).

## The style model

`Style` exists on cells, ranges, rows, columns, row and column collections, and worksheets,
and is typed `IXlsxStyle`. It groups `Font`, `Fill`, `Border`, `Alignment`, `NumberFormat`,
and `Protection`, each an API object that writes through to the cells the style came from.

A cell without an explicit format inherits from its row, then its column, then the worksheet,
then the workbook default. The row wins over the column, so a row format overrides a column
format on the same cell.

```csharp
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Budget");

// Chained setters: SetBold() returns the style, so Fill is reached in the same expression.
sheet.Range("A1:D1").Style.Font.SetBold().Fill.SetBackgroundColor(XlsxColor.FromHtml("#1F4E79"));

// A local variable keeps the connection to A1:D1; later mutations still land on those cells.
IXlsxStyle header = sheet.Range("A1:D1").Style;
header.Font.FontColor = XlsxColor.White;
header.Alignment.Horizontal = XlsxAlignmentHorizontalValues.Center;

sheet.Range("F1:H1").Style = header; // Copies the format; the ranges stay unlinked.

// Worksheet default, used where a cell, row, and column have no format of their own.
sheet.Style.Font.FontName = "Arial";

// A column format is inherited by cells in that column that have no format.
sheet.Column(2).Style.NumberFormat.Format = "#,##0.00";

sheet.Cell("B2").Value = 1234.5;
workbook.SaveAs("budget.xlsx");
```

- There is no `StyleKey` or `IXlsxCellFormat` type in the public API of Aid.Xlsx 1.x and 2.x, so a
  style has no addressable identity. `IXlsxStyle` implements `IEquatable<IXlsxStyle>`, so
  compare style objects to test whether two containers share a format.
- Every setter computes a modified cell format and registers it in the workbook style table.
  Registration is deduplicated, so one property combination across many cells yields one style
  record, setting a property to its current value changes nothing, and assigning `Style` copies
  the format value without linking the two containers.

## Fonts

| Enum | Members |
|---|---|
| `XlsxFontUnderlineValues` | `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting` |
| `XlsxFontVerticalTextAlignmentValues` | `Baseline`, `Subscript`, `Superscript` |
| `XlsxFontScheme` | `None`, `Major`, `Minor` |

```csharp
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Text");

var title = sheet.Cell("A1").Style;
title.Font.FontName = "Calibri";
title.Font.FontSize = 16;
title.Font.Bold = true;
title.Font.FontColor = XlsxColor.FromArgb(0x1F, 0x4E, 0x79);

var note = sheet.Cell("A2").Style;
note.Font.Italic = true;
note.Font.Underline = XlsxFontUnderlineValues.Single;
note.Font.FontScheme = XlsxFontScheme.Minor;

sheet.Cell("A3").Style.Font.VerticalAlignment = XlsxFontVerticalTextAlignmentValues.Superscript;

workbook.SaveAs("text.xlsx");
```

- Font size is stored in twips; the allowed range is 1 to 409.55 points. Out-of-range values
  throw `ArgumentOutOfRangeException`, and sizes are rounded to 1/20 of a point. Font names
  must be non-empty and at most 31 characters (`ArgumentException`), and names that differ
  only in case are equal.
- `FontColor` accepts named colors, `XlsxColor.FromArgb`, `XlsxColor.FromHtml`, indexed colors,
  and theme colors. Reading `.Color` on a theme or automatic color throws
  `InvalidOperationException`; check `ColorType` first.
- `Major` and `Minor` mark the font as a theme heading or body font; Excel resolves the
  typeface from the theme, so the stored name is a hint. `Strikethrough` is a plain `bool`.

## Fills

A solid fill is one color; hatch patterns draw the pattern (foreground) color over the
background color.

| `XlsxFillPatternValues` | Meaning |
|---|---|
| `None`, `Solid` | No fill, single color |
| `LightGray`, `MediumGray`, `DarkGray`, `Gray125`, `Gray0625` | Shading steps |
| `LightGrid`, `DarkGrid`, `LightTrellis`, `DarkTrellis` | Grid and trellis |
| `LightHorizontal`, `DarkHorizontal`, `LightVertical`, `DarkVertical` | Stripes |

```csharp
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Bands");

// Background color alone is enough; the pattern switches from None to Solid.
sheet.Range("A1:B1").Style.Fill.BackgroundColor = XlsxColor.FromHtml("#DDEBF7");

// A hatch needs a pattern, a foreground color, and a background color.
var hatch = sheet.Range("A3:B3").Style.Fill;
hatch.PatternType = XlsxFillPatternValues.LightGrid;
hatch.PatternColor = XlsxColor.Red;
hatch.BackgroundColor = XlsxColor.White;

workbook.SaveAs("bands.xlsx");
```

- Setting a real color while the pattern is `None` switches it to `Solid`, which is why the
  first example needs no `PatternType`; setting `BackgroundColor` to `XlsxColor.Automatic` on
  a solid fill clears it to `None`.
- Setting `PatternType` on an unfilled cell first paints the background with the theme text
  color, so set both colors after choosing the pattern.
- Colors are `XlsxColor` values everywhere: named colors such as `XlsxColor.LightBlue`,
  `XlsxColor.FromArgb`, `XlsxColor.FromHtml("#RRGGBB")`, `XlsxColor.FromIndex`,
  `XlsxColor.FromTheme`, plus `XlsxColor.Automatic` and `XlsxColor.NoColor`. `workbook.Theme`
  exposes the theme palette and `ResolveThemeColor` turns a `XlsxThemeColor` into an RGB value.
- Gradient fills are not exposed by the public API; `IXlsxFill` is pattern-only.

## Borders

A border has left, top, right, bottom, and diagonal slots, plus `InsideBorder` and
`OutsideBorder` for multi-cell ranges. The inside and outside members are set-only and apply
to every cell of the range so that the lines between cells are drawn.

| `XlsxBorderStyleValues` | Members |
|---|---|
| Line weights | `Hair`, `Thin`, `Medium`, `Thick`, `Double` |
| Dashed | `Dashed`, `Dotted`, `DashDot`, `DashDotDot`, `SlantDashDot` |
| Medium dashed | `MediumDashed`, `MediumDashDot`, `MediumDashDotDot` |

```csharp
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Borders");

sheet.Range("A1:C4").Style.Border.SetOutsideBorder(XlsxBorderStyleValues.Thin);
sheet.Range("A1:C4").Style.Border.SetInsideBorder(XlsxBorderStyleValues.Hair);

// Style first, then color: a color on an invisible border line is discarded.
var total = sheet.Cell("C4").Style.Border;
total.TopBorder = XlsxBorderStyleValues.Double;
total.TopBorderColor = XlsxColor.Red;

var struck = sheet.Cell("A6").Style.Border;
struck.DiagonalBorder = XlsxBorderStyleValues.Thin;
struck.DiagonalUp = true;

workbook.SaveAs("borders.xlsx");
```

- A border line whose style is `None` cannot hold a color. Assign the style before the color,
  otherwise the color is dropped and the line later appears in the default color.
- `DiagonalUp` and `DiagonalDown` choose the direction and both need a visible
  `DiagonalBorder`. Range borders are stored per cell, so an outside border on a large range
  materializes formats for the edge cells only.

## Alignment

| Horizontal (`XlsxAlignmentHorizontalValues`) | Vertical (`XlsxAlignmentVerticalValues`) |
|---|---|
| `General`, `Left`, `Center`, `Right` | `Bottom`, `Center`, `Top` |
| `Fill`, `Justify`, `CenterContinuous`, `Distributed` | `Justify`, `Distributed` |

`XlsxAlignmentReadingOrderValues` selects the text direction inside the cell:
`ContextDependent`, `LeftToRight`, `RightToLeft`.

```csharp
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Alignment");

var wrapped = sheet.Range("A1:B1").Style.Alignment;
wrapped.WrapText = true;
wrapped.Vertical = XlsxAlignmentVerticalValues.Center;

var indented = sheet.Cell("A3").Style.Alignment;
indented.Horizontal = XlsxAlignmentHorizontalValues.Left;
indented.Indent = 2;

var shaped = sheet.Cell("B3").Style.Alignment;
shaped.TextRotation = 45;
shaped.ShrinkToFit = true;
shaped.ReadingOrder = XlsxAlignmentReadingOrderValues.RightToLeft;

sheet.Row(1).AdjustToContents();
workbook.SaveAs("alignment.xlsx");
```

- `Indent` is valid only with `Left`, `Right`, or `Distributed` alignment. With `General` a
  positive indent first switches the alignment to `Left`; any other alignment throws
  `InvalidOperationException`.
- `TextRotation` accepts `-90` to `90` degrees, or `255` for stacked vertical text; use
  `TopToBottom = true` instead of passing `255`.
- `WrapText` does not grow the row. Call `AdjustToContents()` on the row after enabling it,
  otherwise the text is clipped at the current height.

## Number formats

`NumberFormat.Format` takes an Excel format code. The code is stored verbatim and is locale
independent: separators in the code are placeholders, and the culture passed to
`GetFormattedString` decides what is rendered.

```csharp
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Formats");

sheet.Cell("A1").Value = 1234.5;
sheet.Cell("A1").Style.NumberFormat.Format = "#,##0.00"; // Thousands separator.

sheet.Cell("A2").Value = 45200;
sheet.Cell("A2").Style.NumberFormat.Format = "yyyy-mm-dd"; // Date code on a serial value.

sheet.Cell("A3").Value = "00123";
sheet.Cell("A3").Style.NumberFormat.Format = "@"; // Text, keeps leading zeros.

// Predefined formats are addressable by id through the nested enums.
sheet.Cell("A4").Value = 0.1875;
sheet.Cell("A4").Style.NumberFormat.NumberFormatId =
    (int)XlsxPredefinedFormat.Number.PercentPrecision2;

// Up to four sections: positive; negative; zero; text.
sheet.Cell("A5").Value = -42;
sheet.Cell("A5").Style.NumberFormat.Format = "#,##0;[Red]-#,##0;\"-\";@";

workbook.SaveAs("formats.xlsx");
```

- `NumberFormatId` accepts only predefined ids: members of `XlsxPredefinedFormat.Number` or
  `XlsxPredefinedFormat.DateTime` cast to `int`, plus `XlsxPredefinedFormat.General`. Any
  other value throws `ArgumentOutOfRangeException`; use `Format` for custom codes.
- Both `Text` members map to `@`. The enums also cover percentages, scientific notation, and
  fractions, and date and time codes use `yyyy`, `mm`, `dd`, `hh`, `ss`, and `[h]` for
  elapsed hours. Excel decides whether `m` means month or minute from its position.
- An invalid or unsupported code is accepted when assigned; rendering then falls back to the
  code text itself, so verify custom codes against Excel. `DateFormat` returns the same object
  as `NumberFormat`, and `GetFormattedString(culture)` controls decimal and group separators
  at render time without changing the stored code. See
  [Data types and conversions](data-types-and-conversions.md) for parsing and converting
  values.

## Cell protection

`Locked` (default `true`) and `Hidden` (default `false`) take effect only once the worksheet
is protected, and `Hidden` hides the formula rather than the value. See
[Protection](protection.md) for sheet protection, allowed elements, and passwords.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Input");

sheet.Cell("B2").Style.Protection.Locked = false;
sheet.Cell("C2").Style.Protection.Locked = false;

sheet.Protect();
workbook.SaveAs("input.xlsx");
```

Unlock the cells the user may edit before protecting the sheet.

## Rows and columns

`Width` and `Height` use Excel units: column width in characters of the default font, row
height in points. A new workbook starts at 8.43 characters and 15 points; sheet-wide defaults
are `IXlsxWorksheet.ColumnWidth` and `IXlsxWorksheet.RowHeight`.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Layout");

sheet.Cell("A1").Value = "A long description that needs room";
sheet.Column(1).Width = 42;
sheet.Row(1).Height = 30;

// Fit to content within a row span and width limits.
sheet.Columns(2, 4).AdjustToContents(1, 100, 8, 40);
sheet.Column(6).Hide();
sheet.Row(6).Hide();

workbook.SaveAs("layout.xlsx");
```

- `AdjustToContents` measures text with the workbook graphic engine, so results depend on the
  cell font metrics. Overloads exist for a row or column span and for minimum and maximum size.
- `Row.ClearHeight()` restores the sheet default height and clears the custom-height flag;
  there is no equivalent for columns, assign `Width` instead.
- `IsHidden` is read-only; use `Hide()` and `Unhide()`. Hiding is independent of grouping.

### Grouping and outlines

Outline levels 0 to 8 and collapse state let detail rows fold behind one expander.

```csharp
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Outline");

sheet.Rows(3, 6).Group();
sheet.Rows(7, 10).Group(2);
sheet.Columns(2, 3).Group(1, collapse: true);

// Collapse or expand by level for the whole sheet.
sheet.CollapseRows(1);
sheet.ExpandColumns();

sheet.Outline.SummaryVLocation = XlsxOutlineSummaryVLocation.Top;
sheet.Outline.SummaryHLocation = XlsxOutlineSummaryHLocation.Left;

workbook.SaveAs("outline.xlsx");
```

- `OutlineLevel` must be 0 to 8, otherwise `ArgumentOutOfRangeException` is thrown; `Group()`
  increases the level by one and `Group(level)` sets it. `CollapseRows(level)` and
  `CollapseColumns(level)` accept a level from 1 to 8 and affect rows or columns with exactly
  that level.
- `SummaryVLocation` and `SummaryHLocation` place the subtotal row or column above or below,
  left or right of the detail; the defaults are `Bottom` and `Right`.

## Worksheet features

Sheets are created with `workbook.Worksheets.Add(name)`. Names are 1 to 31 characters, must
not contain `: \ / ? * [ ]`, must not start or end with an apostrophe, and must be unique
ignoring case; violations throw `ArgumentException`.

```csharp
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

using var workbook = new XlsxWorkbook();
var summary = workbook.Worksheets.Add("Summary");
var details = workbook.Worksheets.Add("Details");
var notes = workbook.Worksheets.Add("Notes");

summary.SheetView.FreezeRows(1);
summary.SheetView.SplitRow = 3;
summary.SheetView.ZoomScale = 125;
summary.ColumnWidth = 12; // Sheet default for columns without an explicit width.
summary.RowHeight = 18; // Sheet default for rows without an explicit height.

details.SheetView.Freeze(2, 1);
details.SetTabColor(XlsxColor.FromArgb(0x2E, 0x75, 0xB6));
details.Position = 1;
notes.Visibility = XlsxWorksheetVisibility.Hidden;

var copy = details.CopyTo("Details copy", 2);
copy.SetTabColor(XlsxColor.Lime);

workbook.SaveAs("worksheets.xlsx");
```

- `Visibility` supports `Visible`, `Hidden`, and `VeryHidden`. Hiding a sheet also clears its
  tab selection. `SetTabColor` and `TabColor` accept any color, including theme colors.
- `Position` is 1-based and moves the sheet; it must not exceed the sheet count plus one.
  `Add(name, position)` inserts at a position and shifts the following sheets down.
- `CopyTo` copies values, styles, row and column settings, data validation, conditional
  formatting, tables, defined names, merged ranges, pictures, page setup, and the sheet view.
  It does not copy comments.
- `sheet.Delete()` and `workbook.Worksheets.Delete(name)` remove the sheet and its part on
  save; a deleted sheet cannot be copied.
- `Freeze`, `FreezeRows`, and `FreezeColumns` add frozen panes. `SplitRow` and `SplitColumn`
  set the same split without choosing the frozen form: Aid.Xlsx always writes frozen panes, so
  a free split that scrolls on both sides is not available.
- `ZoomScale` is clamped to 10 to 400 and omitted at 100. The view mode is set with
  `SheetView.SetView`: `XlsxSheetViewOptions.Normal`, `PageBreakPreview`, or `PageLayout`.
  `ShowGridLines`, `ShowRowColHeaders`, `ShowFormulas`, `ShowZeros`, and `RightToLeft` are
  sheet-wide display flags; `Alignment.ReadingOrder` overrides `RightToLeft` per cell.

## Page setup and printing

`IXlsxPageSetup` holds orientation, paper size, margins, print area, repeating titles,
fit-to-page settings, header and footer content, and manual page breaks.

```csharp
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");
sheet.Cell("A1").Value = "Region";
sheet.Cell("B1").Value = "Revenue";
sheet.Range("A2:B200").Value = 1;

var setup = sheet.PageSetup;
setup.PageOrientation = XlsxPageOrientation.Landscape;
setup.SetPaperSize(XlsxPaperSize.A4Paper);
setup.Margins.Left = 0.5;
setup.PrintAreas.Add("A1:B200");
setup.SetRowsToRepeatAtTop(1, 1);
setup.SetColumnsToRepeatAtLeft("A", "A");
setup.FitToPages(1, 3);

setup.Header.Center.AddText("Revenue by region", XlsxHFOccurrence.AllPages);
setup.Header.Center.AddText(XlsxHFPredefinedText.PageNumber, XlsxHFOccurrence.AllPages);
setup.Footer.Right.AddText("Confidential", XlsxHFOccurrence.AllPages);

setup.AddHorizontalPageBreak(100);
setup.AddVerticalPageBreak(2);

workbook.SaveAs("report.xlsx");
```

| Enum | Members |
|---|---|
| `XlsxPageOrientation` | `Default`, `Portrait`, `Landscape` |
| `XlsxPageOrderValues` | `DownThenOver`, `OverThenDown` |
| `XlsxHFOccurrence` | `AllPages`, `OddPages`, `EvenPages`, `FirstPage` |
| `XlsxHFPredefinedText` | `PageNumber`, `NumberOfPages`, `Date`, `Time` |
| `XlsxHFPredefinedText` | `Path`, `FullPath`, `File`, `SheetName` |
| `XlsxShowCommentsValues` | `None`, `AtEnd`, `AsDisplayed` |
| `XlsxPrintErrorValues` | `Blank`, `Dash`, `Displayed`, `NA` |

- `Scale` and fit-to-page are mutually exclusive. `FitToPages(wide, tall)` clears the scale,
  and setting `Scale` or calling `AdjustTo` clears the page counts. At least one non-zero page
  count is required for a fit-to-page setting to be written.
- The plain-text `AddText` overload on a header or footer requires an occurrence; only the
  predefined-text overload defaults to `AllPages`. Combined text is limited to 255 characters
  including markup, otherwise `ArgumentOutOfRangeException` is thrown. `AddImage` is declared
  but throws `NotImplementedException`, so header logos cannot be added.
- `SetRowsToRepeatAtTop` and `SetColumnsToRepeatAtLeft` accept row indexes or column letters;
  the first value must be positive and not greater than the last. Page breaks are stored as
  ids, deduplicated, and sorted, and `RowBreaks` and `ColumnBreaks` expose the lists.
- `ShowGridlines`, `ShowRowAndColumnHeadings`, `CenterHorizontally`, and `CenterVertically`
  map to the print options. `FirstPageNumber` is nullable and enables an explicit first page
  number, and `SetBlackAndWhite()`, `SetDraftQuality()`, `SetShowComments(...)`, and
  `SetPrintErrorValue(...)` map to the remaining options.
