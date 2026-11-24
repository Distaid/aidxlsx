# Hyperlinks and sparklines

A hyperlink turns a cell into navigation to a web page, a file, an email address, or
another cell. A sparkline is a small chart drawn inside a cell that summarises a row
or column of values without adding a chart object. Both are stored on the worksheet,
so they move with their cells and survive a load and save.

## Hyperlinks

Use a hyperlink when a cell should take the reader somewhere else. The cell text is
the visible link; the link object holds the target and tooltip. `SetHyperlink` attaches
a link to one cell.

### External links

A target that is an absolute URI or a string beginning with `.` is stored as an
external relationship. The second constructor argument is the tooltip.

```csharp
using System;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");

var cell = sheet.Cell("A2");
cell.Value = "Open the vendor portal";
cell.SetHyperlink(new XlsxHyperlink("https://example.com/vendors", "Vendor portal"));

workbook.SaveAs("report.xlsx");
```

### Internal links

A target that is not an absolute URI, and does not begin with `.`, is an internal
address. The constructors that take a cell or a range build the address from that
object, including the sheet name.

```csharp
using System;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var summary = workbook.Worksheets.Add("Summary");
var details = workbook.Worksheets.Add("Details");
details.Cell("A1").Value = "Details";

summary.Cell("B2").Value = "Jump to Details";
summary.Cell("B2").SetHyperlink(
    new XlsxHyperlink(details.Cell("A1"), "Go to the details sheet"));
```

A literal address such as `new XlsxHyperlink("Details!A5")` produces the same target.

### Email and relative file targets

Use a `Uri` for targets that are not plain web addresses. `mailto:` produces an email
link; a relative `Uri` produces a link resolved from the workbook location.

```csharp
using System;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Contacts");

sheet.Cell("A2").Value = "Email support";
sheet.Cell("A2").SetHyperlink(
    new XlsxHyperlink(new Uri("mailto:support@example.com?subject=Question")));

sheet.Cell("A3").Value = "Open the shared tracker";
sheet.Cell("A3").SetHyperlink(
    new XlsxHyperlink(new Uri("../shared/tracker.xlsx", UriKind.Relative)));
```

The string overload treats only strings that start with `.` as relative external
targets. A string such as `reports/q1.xlsx` is not an absolute URI, so it becomes an
internal address instead. Use the `Uri` overload for relative file links.

### Link kind

`IsExternal` reports the kind. `ExternalAddress` returns the `Uri` for external links
and `null` for internal ones. `InternalAddress` returns the escaped sheet and cell
reference for internal links and `null` for external ones.

### Tooltip and display text

`Tooltip` is the hover text; an empty tooltip is not written. There is no separate
display-text property: the cell value is what the reader sees. Change the value to
change the link text, and format the cell if the link should fit a column.

`SetHyperlink` applies the theme hyperlink colour and underline only when the cell
still uses the worksheet default style. If the cell already has its own format, set
`Style.Font.FontColor` and `Style.Font.Underline` yourself.

### Reading links

`HasHyperlink` is a pure check. `GetHyperlink` returns the existing link but creates
a new, target-less link when there is none, so do not use it as a probe.

```csharp
using System;
using Aid.Xlsx.Model.Hyperlinks;

var cell = sheet.Cell("A2");
if (cell.HasHyperlink)
{
    XlsxHyperlink link = cell.GetHyperlink();
    string target = link.IsExternal
        ? link.ExternalAddress.ToString()
        : link.InternalAddress;
    Console.WriteLine($"{target} tooltip={link.Tooltip}");
}
```

The worksheet exposes the full collection through `Hyperlinks`. `Get` throws
`KeyNotFoundException` when the cell has no link, `TryGet` is the non-throwing form,
and enumeration returns links in the order they were loaded or added.

### Updating and removing links

Change `Tooltip`, `ExternalAddress`, or `InternalAddress` on the existing object.
Setting `ExternalAddress` makes the link external; setting `InternalAddress` makes it
internal. Assigning the property again replaces the target in place.

```csharp
using System;
using Aid.Xlsx.Model.Hyperlinks;

XlsxHyperlink link = sheet.Cell("A2").GetHyperlink();
link.Tooltip = "Updated tooltip";
link.ExternalAddress = new Uri("https://example.com/v2/vendors");

bool removed = sheet.Hyperlinks.Delete(sheet.Cell("A2").Address);
```

There is no `RemoveHyperlink` method on the cell. Use one of these three forms:

| Call | Effect |
| --- | --- |
| `cell.SetHyperlink(null)` | Removes the link, keeps the cell formatting |
| `link.Delete()` | Removes the link and resets hyperlink styling |
| `sheet.Hyperlinks.Delete(address)` | Removes the link by address, same style reset |

`SetHyperlink(null)` does not clear the blue underline, while the two delete forms do.
Deleting always resets the underline to the worksheet font underline, and resets the
font colour only when the cell still uses the theme hyperlink colour; a custom font
colour is left alone.

### Edge cases and limits

- A link belongs to exactly one cell. When a file stores one hyperlink over a range,
  loading creates a separate link for every covered cell, and saving writes them back
  as separate cell links. A multi-cell anchor is therefore not preserved as one object.
  A link can still target a range; only its anchor is limited to one cell.
- Renaming the target worksheet does not rewrite internal addresses, and deleting the
  target sheet leaves a stale address. The link's own cell moves with row and column
  insertion and deletion, but the target text is not transformed.
- `GetHyperlink` and `CreateHyperlink` create a link without a target. A link with
  neither address has nothing to save; configure the target before saving.
- Links are not validated. A malformed or unreachable target is written as given and
  reported by Excel when the link is clicked.
- Links are stored per worksheet; copying a worksheet copies its links. There is no
  workbook-wide link collection.

## Sparklines

Use a sparkline when a cell should show a trend next to a value, for example monthly
sales per product. Data and location are set per sparkline; appearance is set per
group. A group can contain many sparklines and is the unit that is written to the file.

### Create a sparkline group

This example writes six months of values and draws a column sparkline in `D1`.

```csharp
using System;
using Aid.Xlsx.Model.Sparklines;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Trends");

string[] months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
double[] sales = [120, 135, 128, 160, 172, 190];
for (var i = 0; i < months.Length; i++)
{
    sheet.Cell(i + 1, 1).Value = months[i];
    sheet.Cell(i + 1, 2).Value = sales[i];
}

IXlsxSparklineGroup group = sheet.SparklineGroups.Add("D1", "B1:B6");
group.SetType(XlsxSparklineType.Column)
    .SetShowMarkers(XlsxSparklineMarkers.HighPoint | XlsxSparklineMarkers.LowPoint);
group.Style = XlsxSparklineTheme.Accent2;
group.Style.SeriesColor = XlsxColor.FromHtml("FF1F77B4");
group.LineWeight = 1.25;

workbook.SaveAs("trends.xlsx");
```

`Add(location, sourceData)` creates a group with one sparkline. The address overload
takes A1 addresses; the range overloads take range objects. `Add(group)` attaches an
existing group that already belongs to the same worksheet. `Add` on the worksheet
collection returns the group, so set group properties immediately.

### One sparkline per row or column

Adding a location range creates one sparkline per cell. A single-row location requires
a source range of the same width; a single-column location requires the same height.
Each sparkline then reads the matching row or column.

```csharp
// Locations D2:D4, one per row; each reads A:C of its own row.
IXlsxSparklineGroup rowGroup =
    sheet.SparklineGroups.Add(sheet.Range("D2:D4"), sheet.Range("A2:C4"));
```

A single-column location works the same way: `D10:D13` with `A10:A13` reads one
source row per location. A multi-row, multi-column location range throws
`ArgumentException`, as does a width or height mismatch between the two ranges.

### Group and sparkline properties

A group owns everything visual: `Type`, `Style`, `LineWeight`, `ShowMarkers`,
`DisplayEmptyCellsAs`, `DisplayHidden`, `DateRange`, `HorizontalAxis`, `VerticalAxis`,
and the `CopyFrom` and `CopyTo` helpers. A sparkline owns only `Location` and
`SourceData`; there are no per-sparkline colours or marker flags. Moving a sparkline to
a cell that already has one replaces the existing sparkline.

### Types and markers

| `XlsxSparklineType` | Appearance |
| --- | --- |
| `Line` (default) | Line chart |
| `Column` | Columns above and below the baseline |
| `Stacked` | Stacked columns; this is the win/loss rendering |

The API names the win/loss type `Stacked`; there is no separate `WinLoss` member. Set
`XlsxSparklineType.Stacked` to get that rendering.

`ShowMarkers` is a flags enum. `Markers` marks every point; the other flags mark
specific points. `All` combines the categories.

| Flag | Value | Marks |
| --- | --- | --- |
| `None` | 0 | No markers |
| `HighPoint` | 2 | Highest value |
| `LowPoint` | 4 | Lowest value |
| `FirstPoint` | 8 | First value |
| `LastPoint` | 16 | Last value |
| `NegativePoints` | 32 | Values below zero |
| `Markers` | 64 | Every point |
| `All` | 126 | Every category above |

### Colours and style

A style holds seven colours: `SeriesColor`, `NegativeColor`, `MarkersColor`,
`HighMarkerColor`, `LowMarkerColor`, `FirstMarkerColor`, and `LastMarkerColor`. Assign
a `XlsxSparklineTheme` style and override individual colours. Themes cover `Accent1` to
`Accent6` with `Lighter40`, `Darker25`, and `Darker50` variants, `Colorful1` to
`Colorful6`, and `Dark1` to `Dark6`; `Default` is `Dark5`.

```csharp
group.Style = XlsxSparklineTheme.Colorful3;
group.Style.SeriesColor = XlsxColor.FromHtml("FF2E75B6");
group.Style.NegativeColor = XlsxColor.Red;
group.Style.MarkersColor = XlsxColor.Black;
group.Style.SetHighMarkerColor(XlsxColor.Green);
```

Assigning `group.Style = null` throws `ArgumentNullException`. Theme getters return a
new style instance each time, so assigning a theme and then editing the group style is
safe.

### Axes

The horizontal axis has `Color`, `IsVisible`, `RightToLeft`, and the read-only
`DateAxis`. `DateAxis` is `true` exactly when `DateRange` is set. The vertical axis has
`MinAxisType` and `MaxAxisType` (`Automatic`, `SameForAll`, or `Custom`) plus
`ManualMin` and `ManualMax`.

Setting a manual minimum or maximum switches that axis type to `Custom`. Switching the
type away from `Custom` clears the corresponding manual value. A date range must be one
row high or one column wide; its length is not checked against the data range.

```csharp
group.DateRange = sheet.Range("A2:A7");
group.HorizontalAxis.SetVisible(true).SetColor(XlsxColor.FromHtml("FF808080"));
group.VerticalAxis.SetMinAxisType(XlsxSparklineAxisMinMax.SameForAll);
group.VerticalAxis.SetManualMax(200);
```

### Reading and removing sparklines

`cell.HasSparkline` and `cell.Sparkline` find the sparkline in a cell. The group
collection can also search a range. `SourceData` is `null` when the stored formula no
longer resolves to a range in the workbook, for example after the source was deleted.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Sparklines;

if (sheet.Cell("D1").HasSparkline)
{
    IXlsxSparkline spark = sheet.Cell("D1").Sparkline;
    Console.WriteLine(spark.SourceData?.RangeAddress.ToStringRelative());
    spark.SetSourceData(sheet.Range("B2:B7"));
    spark.SetLocation(sheet.Cell("E1"));
}

foreach (IXlsxSparkline spark in sheet.SparklineGroups.GetSparklines(sheet.Range("D1:D10")))
    Console.WriteLine(spark.Location.Address.ToStringRelative());

sheet.SparklineGroups.Remove(sheet.Range("D1:D10"));
sheet.Cell("E1").Clear(XlsxClearOptions.Sparklines);
```

`Remove(cell)` removes one sparkline, `Remove(range)` removes every sparkline anchored
in the range, `Remove(group)` removes a whole group, and `RemoveAll()` removes every
group and sparkline on the worksheet. Removing the last sparkline of a group does not
remove the group object, but the group is not written to the file when it is empty.

### Edge cases and limits

- A cell holds at most one sparkline. Adding to an occupied cell silently removes the
  old sparkline from its group first; moving a sparkline onto an occupied cell does the
  same.
- Source data must be one row or one column, and must belong to the same workbook.
  Other ranges throw `ArgumentException`.
- A group belongs to one worksheet. `CopyTo` throws `InvalidOperationException` when
  the target is the same worksheet; it copies the group, its sparkline locations, and
  its source formulas to the target sheet.
- `DisplayEmptyCellsAs` controls gaps: `Interpolate` spans blanks, `NotPlotted` leaves a
  gap, and `Zero` plots zero. `DisplayHidden` includes hidden rows and columns.
- The file format stores one cell reference per sparkline. When a producer writes
  several references in one entry, loading uses the first one.
- There is no per-point formatting, no data labels, and no axis title. Axis options are
  limited to the members above. Sparklines live in an `x14` worksheet extension; Excel
  reads them, but older consumers may ignore the extension.
- Inserting or deleting rows and columns shifts sparkline locations and adjusts source
  references automatically, including references on other worksheets.
