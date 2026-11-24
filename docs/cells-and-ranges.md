# Cells and ranges

A worksheet is a fixed grid of 1,048,576 rows and 16,384 columns. Address one position
with a cell, a rectangular block with a range, and a whole row, column, or sheet with a
special range. Rows and columns are 1-based: row `1`, column `1` is `A1`.

## Address a cell

Use a cell when you read or write one position. `Cell` accepts A1 text, row and column
numbers, or a column letter. The returned `IXlsxCell` is a live reference to the grid.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Addressing");
IXlsxCell a1 = sheet.Cell("A1");
IXlsxCell b2 = sheet.Cell(2, 2);
IXlsxCell c3 = sheet.Cell(3, "C");
a1.Value = "Top left";
b2.Value = 42;
c3.Value = true;

Console.WriteLine($"{b2.Address}: row {b2.Address.RowNumber}, " +
                  $"column {b2.Address.ColumnLetter}");
```

| Call | Address |
| --- | --- |
| `sheet.Cell("A1")` | A1 notation; `$` markers are accepted. |
| `sheet.Cell(row, column)` | Numeric coordinates, starting at 1. |
| `sheet.Cell(row, "D")` | Mixed row number and column letter. |
| `sheet.Cell(IXlsxAddress)` | An address obtained from another cell or range. |

- A string that is not an A1 address is resolved as a defined name; if neither matches,
  the call throws `ArgumentException`.
- A numeric row or column outside `1..1048576` or `1..16384` throws
  `ArgumentOutOfRangeException`. Use `XlsxHelper.MaxRowNumber` and
  `XlsxHelper.MaxColumnNumber` in validation code.
- R1C1 text is not a cell input format. R1C1 is available for formulas and for address
  formatting.
- On a merged range, only the upper-left cell holds the value. Writing through an inferior
  cell is ignored.
- Navigate with `CellAbove()`, `CellBelow()`, `CellLeft()`, and `CellRight()`, each with a
  step overload; `range.Relative(source, target)` maps a range between two anchors.

## Address a range

Use a range for a rectangular block that you format, fill, traverse, or edit as a unit.
The same block can be described with text, numbers, or existing cells.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Data");
IXlsxRange block = sheet.Range("B2:D4");
IXlsxRange full = workbook.RangeFromFullAddress("Data!B2:D4", out _);
Console.WriteLine(full.RowCount());  // 3
IXlsxRange columnB = sheet.Range("B:B");
Console.WriteLine(columnB.RangeAddress.IsEntireColumn());   // True
Console.WriteLine(sheet.AsRange().RangeAddress.IsEntireSheet()); // True

foreach (IXlsxRange part in sheet.Ranges("A1:B2,D4:E5"))
    Console.WriteLine(part.RangeAddress);

Console.WriteLine(sheet.Cell("B2").Address.ToString(XlsxReferenceStyle.R1C1)); // R2C2
workbook.ReferenceStyle = XlsxReferenceStyle.R1C1; // saved as refMode="R1C1"
```

| Call | Result |
| --- | --- |
| `sheet.Range("A1:C3")` | Rectangular block from `A1` to `C3`. |
| `sheet.Range(1, 1, 3, 3)` | The same block with row and column numbers. |
| `sheet.Range("A1", "C3")` | Corner cells as text. |
| `sheet.Range(firstCell, lastCell)` | Corner cells as `IXlsxCell` values. |
| `sheet.Range("B:B")` / `"2:4"` | A whole column or a whole row band. |
| `sheet.AsRange()` | The whole sheet, `A1:XFD1048576`. |
| `sheet.Ranges("A1:B2,D4")` | A collection from comma-separated text. |
| `workbook.RangeFromFullAddress(...)` | Resolves a sheet-qualified address. |

- Pass the upper-left corner first. The library does not reorder corners, so a reversed
  address such as `C3:A1` produces an unusable range.
- `sheet.Range("Other!A1:A2")` parses the address but applies it to the current worksheet.
  Use `workbook.RangeFromFullAddress` when the sheet matters.
- Text that is not an address, range, table reference, or defined name throws
  `ArgumentException`. Out-of-bounds corners throw `ArgumentOutOfRangeException`.
- Range coordinates are relative to the range: `Range("B2:D4").Cell(1, 1)` is `B2` and
  `.Row(2)` is worksheet row 3.
- `XlsxReferenceStyle.R1C1` changes how addresses are formatted and sets `refMode` on save.
  `Default` behaves as A1 for formatting and omits the attribute. Formula text is always
  stored in A1 form on disk.
- `sheet.Rows()` and `sheet.Columns()` enumerate used axes plus explicitly created ones,
  not every position. Prefer `sheet.Row(3)` or `sheet.Column("B")` for one axis.

## Inspect and combine ranges

Use the range address when you need bounds or counts, and intersection or union when two
areas overlap.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");

IXlsxRange range = sheet.Range("B2:D4");
IXlsxRangeAddress address = range.RangeAddress;

Console.WriteLine(address);                        // B2:D4
Console.WriteLine(address.ToStringFixed());        // $B$2:$D$4
Console.WriteLine(address.ToStringRelative(true)); // Report!B2:D4
Console.WriteLine($"{address.RowSpan} x {address.ColumnSpan}");
Console.WriteLine(address.NumberOfCells);          // 9
Console.WriteLine(range.FirstCell().Address);      // B2
Console.WriteLine(range.LastCell().Address);       // D4

IXlsxRangeAddress overlap = range.Intersection(sheet.Range("C3:F6"));
if (overlap.IsValid)
    Console.WriteLine(overlap); // C3:D4

Console.WriteLine(range.Intersects(sheet.Range("D4:D10"))); // True
Console.WriteLine(range.Contains(sheet.Range("C3:C3")));    // True
```

- `NumberOfCells` is an `int`; a full-sheet address exceeds its range, so use `RowSpan`
  and `ColumnSpan` with `long` arithmetic for very large areas.
- `Intersection` returns an invalid address when the ranges do not overlap and throws
  `ArgumentOutOfRangeException` for a range on another worksheet.
- `Union` and `Difference` return cell collections, not ranges, and are empty across
  worksheets.
- `AsRange()` returns `null` for an invalid address; `ToString()` then returns `#REF!`.
- `range.Range(...)` returns a subrange and throws `ArgumentOutOfRangeException` when the
  requested corners fall outside the parent range.

## Traverse cells and used areas

Use `CellsUsed()` for populated positions and `Cells()` on a bounded range for every
position. Use the worksheet-level used-area members to size a report before writing it.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Traversal");

sheet.Cell("A1").Value = "Region";
sheet.Cell("B1").Value = "Units";
sheet.Cell("A2").Value = "North";
sheet.Cell("B2").Value = 120;
sheet.Cell("A4").Value = "South";
sheet.Cell("B4").Value = 80;
sheet.Cell("C4").Value = "";                         // empty text, not blank
sheet.Cell("D4").CreateComment().AddText("checked"); // counts as content

foreach (IXlsxCell cell in sheet.Range("A1:D4").CellsUsed())
    Console.WriteLine($"{cell.Address}: {cell.GetFormattedString()}");

Console.WriteLine(sheet.FirstRowUsed()!.RowNumber()); // 1
Console.WriteLine(sheet.LastRowUsed()!.RowNumber());  // 4
Console.WriteLine(sheet.RangeUsed()!.RangeAddress);    // A1:D4

foreach (IXlsxCell cell in sheet.Range("A1:B2").Cells())
    Console.WriteLine(cell.Address);
```

| Member | Returns |
| --- | --- |
| `Cells()` | Every position in the range, in row-major order. |
| `CellsUsed()` | Positions that are not empty under `AllContents`. |
| `CellsUsed(options)` | Positions under the selected flags. |
| `Cells(bool usedCellsOnly)` | Bounded enumeration with or without empty positions. |
| `FirstCellUsed()` / `LastCellUsed()` | Topmost-leftmost and bottommost-rightmost used cell. |
| `RangeUsed()` | Smallest range containing the used cells, or `null`. |
| `Worksheet.FirstColumnUsed()` / `LastColumnUsed()` | Bounds of used columns, or `null`. |
| `Worksheet.RowsUsed()` / `ColumnsUsed()` | Used axes, not every axis. |

Flags in `XlsxCellsUsedOptions` are additive. `AllContents` is
`Contents | DataType | Comments`, so a note alone makes a cell used. `NormalFormats` adds
cells whose format differs from inherited formatting; `All` adds every modeled feature.
`Cells(false)` on a whole column or sheet is effectively unbounded and should be avoided.

A blank cell has `DataType == Blank`. A cell assigned an empty string has
`DataType == Text`, `GetText()` returns `""`, and `IsEmpty()` returns `true`. Both are
equivalent for used-cell traversal but differ for typed reads. A formula that returns an
empty string is not empty because `HasFormula` is true.

## Edit ranges

Use insertion and deletion when a block grows or shrinks, and `Clear`, `CopyTo`, or
`Transpose` to rearrange existing content. Range insertions shift only the columns or rows
covered by the range; row and column objects shift the whole axis.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Edits");

sheet.Cell("A1").Value = "Header";
sheet.Cell("A2").Value = 10;
sheet.Cell("A3").Value = 20;

sheet.Range("A2:C2").InsertRowsAbove(1); // shift columns A:C only
sheet.Row(1).InsertRowsAbove(1);         // shift every column
sheet.Column("B").InsertColumnsBefore(2); // shift every row
sheet.Range("A2:C2").Delete(XlsxShiftDeletedCells.ShiftCellsUp);

IXlsxRange copy = sheet.Range("A1:B2").CopyTo(sheet.Cell("E1"));
Console.WriteLine(copy.RangeAddress); // E1:F2

sheet.Range("A1:B2").Clear(XlsxClearOptions.Contents);      // keep formats
sheet.Range("A1:B2").Clear(XlsxClearOptions.NormalFormats); // keep values

sheet.Range("H1:I3").Transpose(XlsxTransposeOptions.MoveCells);

IXlsxRange block = sheet.Range("A1:B2");
IXlsxRangeBase grown = block.Grow(1);      // A1:C3, clamped at sheet edges
IXlsxRangeBase? smaller = block.Shrink(1); // null when nothing remains
workbook.SaveAs("edits.xlsx");
```

| Call | Behavior |
| --- | --- |
| `range.InsertRowsAbove(n)` / `Below(n)` | Shifts rows down; `expandRange` grows the range. |
| `range.InsertColumnsBefore(n)` / `After(n)` | Shifts columns right. |
| `row.InsertRowsAbove(n)` / `column.InsertColumnsBefore(n)` | Shifts the whole axis. |
| `range.Delete(shift)` | Removes the range, shifting up or left. |
| `cell.InsertCellsAbove(n)` and siblings | Convenience forms returning inserted cells. |
| `Clear(XlsxClearOptions.Contents)` | Clears value, formula, and hyperlink; keeps format. |
| `Clear(XlsxClearOptions.NormalFormats)` | Resets direct formatting; keeps contents. |
| `CopyTo(IXlsxCell)` | Copies to the anchor and returns the affected range. |
| `Transpose(options)` | `MoveCells` shifts neighbors; `ReplaceCells` clears them. |
| `Grow(n)` / `Shrink(n)` | Expands or contracts diagonally within sheet bounds. |

- Insertion rewrites formula references workbook-wide, as described in
  [Formulas](formulas.md); shifted formulas follow the data.
- Inserting inside a partially covered merged range removes that merge.
- `CopyTo` copies values, formulas, styles, comments, hyperlinks, merges, and validation
  rules. Formula text is translated relative to the destination, so `=A1*2` in `B1`
  becomes `=B1*2` in `C1`. The destination is cleared first and truncated at the edge.
- `n` must be at least 1 and no more than the axis limit, otherwise the call throws
  `ArgumentOutOfRangeException`.
- There is no `MoveTo` or `Resize`. Copy to the destination and delete the source, or use
  `Grow`/`Shrink` for bounds changes.

## Merge and unmerge

Use merged cells for headings or one logical value over several columns. Only the
upper-left cell retains a value; merging clears the other cells in the block.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Report");

var title = sheet.Range("A1:D1");
title.Merge();
title.FirstCell().Value = "Quarterly report";

foreach (IXlsxRange merged in sheet.MergedRanges)
    Console.WriteLine(merged.RangeAddress);

var inner = sheet.Cell("B1");
Console.WriteLine(inner.IsMerged());                  // True
Console.WriteLine(inner.MergedRange()!.RangeAddress); // A1:D1

title.Unmerge();
```

- `Merge()` removes intersecting merges, clears all cells except the first, and applies the
  first cell's format across the block. `Merge(checkIntersect: false)` skips the cleanup;
  use it only when the block is known to be empty.
- Writing through an inferior merged cell is ignored. Write through `FirstCell()` or the
  upper-left address.
- `MergedRanges` is an `IXlsxRanges` collection; enumerate it to list every merge.
  `Unmerge()` removes the merge but does not restore cleared values.

## Sort

Use range sorting to order rows by one or more keys, or `SortLeftToRight` to order columns
by one or more rows. Keys are positions inside the sorted range, not worksheet columns.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sorted");

sheet.Cell("A1").Value = "Name";
sheet.Cell("B1").Value = "Units";
sheet.Cell("A2").Value = "Pears";
sheet.Cell("B2").Value = 12;
sheet.Cell("A3").Value = "Apples";
sheet.Cell("B3").Value = 30;
sheet.Cell("A4").Value = "Plums";
sheet.Cell("B4").Value = 5;

var data = sheet.Range("A2:B4");
data.Sort(2, XlsxSortOrder.Descending); // key 2 within the range is column B

data.SortColumns.Clear();
data.SortColumns.Add(2, XlsxSortOrder.Descending);
data.SortColumns.Add(1, XlsxSortOrder.Ascending);
data.Sort();
```

| Call | Behavior |
| --- | --- |
| `range.Sort()` | Sorts rows using configured keys, or every column when none are set. |
| `range.Sort(column, order)` | One key; the number is relative to the range. |
| `range.Sort("B", order)` | One key by column letter. |
| `range.SortLeftToRight(order)` | Sorts columns using row keys. |
| `range.SortRows` / `range.SortColumns` | Key collections with `Add`, `Clear`, `Remove`. |
| `worksheet.Sort()` | Sorts the worksheet's used range. |

- The header row is not detected; exclude it from the range as the example does.
- `ignoreBlanks` defaults to `true` and `matchCase` to `false`; both can be set per key.
- `Sort()` with no configured keys sorts by every column in order, effectively by column 1
  first.
- Sorting swaps cell contents and formats inside the range. Formula text moves unchanged,
  so formulas that relied on their original row keep the old references.
- `SortElements.Remove(n)` removes the n-th key in the collection, not the key targeting
  column n.
- A single axis has its own sort: `row.SortLeftToRight(...)` orders a row's cells left to
  right, and `column.Sort(...)` orders a column's cells top to bottom.

## Size, hide, and validate

Use row and column objects for height, width, and visibility; use `XlsxHelper` to validate
user input before constructing an address. Formatting details live in
[Formatting and worksheet features](formatting-and-worksheets.md).

```csharp
using System;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Layout");
sheet.Column("A").Width = 28;
sheet.Row(1).Height = 22;
sheet.Row(3).Hide();
sheet.Columns("B:D").Hide();

foreach (var row in sheet.Rows(1, 3))
    Console.WriteLine($"{row.RowNumber()}: hidden={row.IsHidden}");

if (XlsxHelper.IsValidA1Address("XFD1048576"))
    sheet.Cell("XFD1048576").Value = "Last cell";

try
{
    sheet.Cell("A1048577");
}
catch (ArgumentException exception)
{
    Console.WriteLine(exception.Message);
}

Console.WriteLine(XlsxHelper.MaxRowNumber);    // 1048576
Console.WriteLine(XlsxHelper.MaxColumnNumber); // 16384
Console.WriteLine(XlsxHelper.LastCell);        // XFD1048576
```

- Row and column objects expose `Height`/`Width`, `IsHidden`, `Hide()`/`Unhide()`,
  `CellCount()`, outline members, `AdjustToContents()`, and `ClearHeight()`.

| Operation | Exception |
| --- | --- |
| `Cell("not an address")` | `ArgumentException`. |
| `Cell(0, 1)` or `Cell(1, 16385)` | `ArgumentOutOfRangeException`. |
| `Range("A0:A1")` or `Range(1, 1, 1048577, 1)` | `ArgumentOutOfRangeException`. |
| `range.Cell(0, 0)` | `ArgumentOutOfRangeException`. |
| `range.Range(...)` outside the parent | `ArgumentOutOfRangeException`. |
| Text longer than 32,767 characters | `ArgumentOutOfRangeException`. |

- Fetching a row or column materializes it and applies its formatting to existing cells;
  avoid looping over all 16,384 columns for large sheets.
- Hidden rows and columns still contain readable cells and still count in used-range
  calculations.
- The string validators check syntax, not worksheet bounds:
  `XlsxHelper.IsValidRangeAddress("A9999999")` returns `true` while `Range` still throws
  for the missing row.
