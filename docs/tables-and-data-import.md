# Tables and data import

An Excel table gives a range a name, a header row, structured references and filters.
This guide creates and edits tables, configures totals and auto filters, and covers the
supported ways to write data: a `DataTable`, explicit rows, reflection over objects, and
the trim-safe generated mapper. See [Cells and ranges](cells-and-ranges.md) for addressing.

## Create a table from a range

Call `CreateTable` on a range that already contains the header row. The header cells
become field names, so make them unique and non-empty. A new table enables the header
row, an auto filter, banded rows and the `TableStyleMedium2` theme by default.

```csharp
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Catalog");

sheet.Cell("B1").Value = "Category";
sheet.Cell("C1").Value = "Item";
sheet.Cell("B2").Value = "Fruit";
sheet.Cell("C2").Value = "Pears";
sheet.Cell("B3").Value = "Vegetables";
sheet.Cell("C3").Value = "Carrots";

var table = sheet.Range("B1:C3").CreateTable("CatalogItems");
table.Theme = XlsxTableTheme.TableStyleMedium9;

workbook.SaveAs("catalog.xlsx");
```

`CreateTable()` without an argument generates the first free name in the `Table1`,
`Table2`, ... series. Table names are case-insensitive, must start with a letter, an
underscore or a backslash, and cannot exceed 255 characters. A duplicate on the same
worksheet raises `ArgumentException`; Excel also expects workbook-wide unique names.

A one-row range is extended by an empty data row, because a table needs a header and at
least one data row. A range that overlaps another table or the worksheet auto filter
raises `InvalidOperationException`. A table with no data row raises `EmptyTableException`
on `SaveAs`, so check `table.DataRange` for `null` before saving.

Look a table up on its worksheet with `sheet.Table("CatalogItems")`, or anywhere in the
workbook with `workbook.Table("CatalogItems")`. The workbook overload compares names case
insensitively by default and accepts a `StringComparison` argument. Themes are the
`XlsxTableTheme` constants, from `TableStyleLight1`–`TableStyleLight21` through the
`Medium` and `Dark` families.

`table.ShowHeaderRow`, `ShowRowStripes`, `ShowColumnStripes`, `EmphasizeFirstColumn`,
`EmphasizeLastColumn` and `ShowAutoFilter` are settable properties. `ShowAutoFilter` only
reports `true` while `ShowHeaderRow` is `true`: without a header row there is nowhere to
render filter buttons.

## Fields, rows and the data range

`DataRange` excludes the header row and the totals row, and is `null` when there are no
data rows. A field exposes its header cell, data cells, totals cell and whole column.

```csharp
var table = sheet.Table("CatalogItems");
IXlsxTableField item = table.Field("Item");
Console.WriteLine(item.Index); // 1, zero-based position in the table
item.DataCells.Style.NumberFormat.Format = "@";
```

`table.Field(int)` selects a field by zero-based index, and `table.Fields` enumerates all
fields. `field.Name = "Product"` renames the field, rewrites the header cell, and updates
totals formulas that reference the field. `HeadersRow()` and `TotalsRow()` return `null`
when the corresponding row is hidden.

### Resize, append and delete

`table.Resize("B1:C10")` accepts a range, a range address, two cells or four row/column
numbers. It keeps the fields and rewrites the header row from the new range; empty header
cells receive generated `Column1`, `Column2`, ... names. Resizing a table without a
header row raises `NotImplementedException`; resizing to another worksheet raises
`InvalidOperationException`.

`AppendData` inserts rows below the last data row and above the totals row. It accepts a
`DataTable`, an `IEnumerable` or an `IEnumerable<T>` and matches values by column
position, not by `DataTable` column name.

```csharp
using System.Data;

var rows = new DataTable();
rows.Columns.Add("Category");
rows.Columns.Add("Item");
rows.Rows.Add("Fruit", "Oranges");
table.AppendData(rows);
```

`AppendData` returns `null` and writes nothing for an empty collection, an empty
`DataTable` or a `string`. `propagateExtraColumns: true` copies the value or formula of
the previous last data row into table columns that received no data, which keeps
calculated columns filled. `ReplaceData` resizes the data area to the new row count and
raises `InvalidOperationException` for empty input.

Remove a table with `sheet.Tables.Remove("CatalogItems")`; the cells stay in place.

## Totals rows

`ShowTotalsRow` inserts a totals row below the data. Each field then chooses an aggregate;
`None` means no aggregate, and `Custom` means the cell holds a formula you write.

```csharp
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");

var source = new[]
{
    new object[] { "Category", "Units", "Price" },
    new object[] { "Fruit", 12, 2.5 },
    new object[] { "Fruit", 8, 3.0 },
};
var table = sheet.Cell("A1").InsertData(source).CreateTable("SalesTable");

table.ShowTotalsRow = true;
table.Field("Category").TotalsRowLabel = "Total";
table.Field("Units").TotalsRowFunction = XlsxTotalsRowFunction.Sum;
table.Field("Price").TotalsRowFunction = XlsxTotalsRowFunction.Average;

workbook.SaveAs("sales-totals.xlsx");
```

Built-in functions write an Excel `SUBTOTAL` formula with a structured reference, for
example `SUBTOTAL(109,SalesTable[Units])`.

| `XlsxTotalsRowFunction` | Written formula |
| --- | --- |
| `Sum`, `Count`, `CountNumbers` | `SUBTOTAL(109/103/102, ...)` |
| `Average`, `Maximum`, `Minimum` | `SUBTOTAL(101/104/105, ...)` |
| `StandardDeviation`, `Variance` | `SUBTOTAL(107/110, ...)` |
| `Custom` / `None` | Your `TotalsRowFormulaA1` / no built-in formula |

`TotalsRowLabel` writes text into the totals cell and marks the field as `None`.
Assigning `table.Field("Price").TotalsRowFormulaA1 = "SUBTOTAL(109,[Price])*1.1"`
marks the field as `Custom`.

`TotalsRow()` returns `null` while `ShowTotalsRow` is `false`. Turning it off deletes the
totals row and its formulas; turning it back on inserts a fresh row.

## Auto filters

A table creates its own filter; a plain range can enable a worksheet filter explicitly.
Both return `IXlsxAutoFilter`.

```csharp
var filter = sheet.Range("A1:C3").SetAutoFilter();
filter.Column(1).AddFilter("Fruit");

filter.Column(1).Clear(); // remove column 1 criteria, keep the filter
// filter.Clear();        // disable the filter and unhide every row
```

Column numbers are one-based and relative to the first column of the filtered range, not
to the worksheet. `AddFilter` adds an allowed value; call it repeatedly on the same column
to allow several values, and filter out rows that match none of them. Filtered rows are
hidden, not deleted: they remain in the file with their values.

`VisibleRows` and `HiddenRows` enumerate the rows in the filter range. After changing cell
values, call `Reapply()` to hide rows again according to the stored criteria.

### Custom criteria and connectors

The direct comparison methods replace the criteria of that column. To combine two
conditions, chain `And` or `Or`; a custom filter accepts at most two conditions.

```csharp
var units = table.AutoFilter.Column(2);

units.EqualOrGreaterThan(10).And.LessThan(100);
units.Contains("Pear");
units.NotEqualTo("Carrots");
```

Available comparisons are `EqualTo`, `NotEqualTo`, `GreaterThan`, `LessThan`,
`EqualOrGreaterThan`, `EqualOrLessThan`, `Between`, `NotBetween`, `BeginsWith`,
`NotBeginsWith`, `EndsWith`, `NotEndsWith`, `Contains` and `NotContains`. String methods
treat `*` and `?` as wildcards.

### Top, average, date and sorted filters

`Top(n, XlsxTopBottomType.Items)` keeps the n largest values, `Bottom(n)` the n smallest,
and `Top(n, XlsxTopBottomType.Percent)` keeps the top n percent. `AboveAverage()` and
`BelowAverage()` keep values relative to the column average, and
`AddDateGroupFilter(date, XlsxDateTimeGrouping.Month)` groups by a date part (`Year`,
`Month`, `Day`, `Hour`, `Minute` or `Second`). The value passed to `Top` or `Bottom` must
be between 1 and 500.

`Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks)` physically reorders the rows,
then reapplies the filter. Sorting is not a view setting: the saved file contains the new
row order.

```csharp
table.AutoFilter.Sort(
    columnToSortBy: 1,
    sortOrder: XlsxSortOrder.Ascending,
    matchCase: false,
    ignoreBlanks: true);
```

## Structured references

The calculation engine understands table names in formulas, so `=SUM(Table1[Units])` is
evaluated without converting it to cell addresses. The keywords `[#Headers]`, `[#Data]`,
`[#Totals]`, `[#All]` and `[#This Row]` are supported, including combinations such as
`[[#Headers],[#Data]]`.

```csharp
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");

var rows = new[]
{
    new object[] { "Item", "Units" },
    new object[] { "Pears", 12 },
    new object[] { "Apples", 8 },
};
var table = sheet.Cell("A1").InsertData(rows).CreateTable("SalesTable");
table.ShowTotalsRow = true;
table.Field("Units").TotalsRowFunction = XlsxTotalsRowFunction.Sum;

sheet.Cell("E1").FormulaA1 = "SUM(SalesTable[Units])";
sheet.Cell("E2").FormulaA1 = "COUNTA(SalesTable[Item])";
workbook.RecalculateAllFormulas();

Console.WriteLine(sheet.Cell("E1").GetDouble()); // 20
workbook.SaveAs("sales-reference.xlsx");
```

`[#Totals]` only resolves when the table has a totals row. An unknown table or column name
produces `#REF!`, and a structured reference to another workbook also produces `#REF!`
because external links are not supported. See [Formulas and recalculation](formulas.md)
for when a calculated result is stored in the file.

## Import a DataTable or explicit rows

`InsertTable` writes the column captions as a header row and creates an Excel table.
`InsertData` writes only values, so write headers yourself when you need them.

```csharp
using System.Data;
using Aid.Xlsx.Model.Workbooks;

var data = new DataTable("Stock");
data.Columns.Add("Item", typeof(string));
data.Columns.Add("Units", typeof(int));
data.Rows.Add("Pears", 12);
data.Rows.Add("Apples", 8);

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Stock");
sheet.Cell("A1").InsertTable(data, "StockItems");
workbook.SaveAs("stock.xlsx");
```

Heading text comes from `DataColumn.Caption`, which defaults to `ColumnName`.
`InsertTable(data, tableName: "StockItems", createTable: false)` writes the cells and
header but keeps the table only in memory, so the saved file contains plain data.
`InsertTable` returns `null` when the `DataTable` is `null` or has no columns.

Explicit rows use the same API. Arrays are written row by row, and `transpose: true`
swaps rows and columns:

```csharp
var rows = new[] { new object[] { "Pears", 12 }, new object[] { "Apples", 8 } };

sheet.Cell("A2").InsertData(rows);
sheet.Cell("E2").InsertTable(rows, "FruitTable");
```

A jagged array has no column names, so `InsertTable` generates `Column1`, `Column2`, ...
Write a header row yourself and call `InsertData`, or use a typed source when the table
needs named fields. `InsertData` returns `null` for a `null` or `string` input; an empty
collection writes nothing but returns a one-cell range at the target cell.

## Import typed objects with reflection

`InsertTable<T>` and `InsertData` accept any `IEnumerable<T>` and inspect public fields
and readable properties, including static ones. `[XlsxColumn]` changes the heading or order
and can exclude a member; `[XlsxIgnore]` excludes a member.

```csharp
using System;
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");

var sales = new[]
{
    new Sale("Pears", 12.50m, new DateTime(2026, 5, 1)),
    new Sale("Apples", 8.00m, null),
};

sheet.Cell("A1").InsertTable(sales, "SalesData");
workbook.SaveAs("sales-objects.xlsx");

internal sealed class Sale(string item, decimal amount, DateTime? bookedAt)
{
    [XlsxColumn(Order = 0)] public string Item { get; } = item;
    [XlsxColumn(Order = 1, Header = "Total")] public decimal Amount { get; } = amount;
    [XlsxColumn(Order = 2)] public DateTime? BookedAt { get; } = bookedAt;
    [XlsxIgnore] public string? Note { get; set; }
}
```

Members without `[XlsxColumn]` use their own name and sort after ordered members. The
reflection path reads `Header`, `Order` and `Ignore`, but not `Format`. `Order` defaults
to `0`, so give annotated members distinct values. A `null` item produces a blank row.
These overloads are annotated with `[RequiresUnreferencedCode]`; on trimmed or NativeAOT
paths use a `DataTable`, explicit rows, or the generated mapper in
[NativeAOT and trimming](native-aot.md).

## Import mapped records with the generator

The `Aid.Xlsx` package contains a source generator. Annotate a class or struct with
`[XlsxTable]`, then call `InsertMappedData` or `InsertMappedTable`. The generated mapper
reads members directly and does not use reflection.

```csharp
using System;
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Orders");

var orders = new[]
{
    new Order("Pears", 12.50m, new DateTime(2026, 5, 1), OrderState.Open, null),
    new Order("Apples", 8.00m, null, OrderState.Shipped, TimeSpan.FromHours(36)),
};

sheet.Cell("A1").InsertMappedTable(orders);
workbook.SaveAs("orders.xlsx");

[XlsxTable("Orders")]
internal sealed record Order(
    [property: XlsxColumn(Order = 0)] string Item,
    [property: XlsxColumn(Order = 1, Header = "Total", Format = "#,##0.00")] decimal Amount,
    [property: XlsxColumn(Order = 2, Header = "Ordered")] DateTime? OrderedAt,
    [property: XlsxColumn(Order = 3, Header = "Duration", Format = "[h]:mm")] TimeSpan? LeadTime,
    [property: XlsxColumn(Order = 4)] OrderState State);

internal enum OrderState { Open, Shipped }
```

`InsertMappedTable(orders)` uses the name from `[XlsxTable]`, and
`InsertMappedTable(orders, "OrderTable")` overrides it. If neither supplies a name, the
table gets a generated `TableN` name. `InsertMappedData` writes a plain range; pass
`addHeadings: true` to include the header row.

The generator requires an accessible, non-generic, non-abstract class or struct. It maps
public instance fields and readable public instance properties, including inherited ones.
No parameterless constructor is required: the mapper reads existing instances and never
constructs the type.

| Diagnostic | Meaning |
| --- | --- |
| `XLSX001` | Duplicate header name or duplicate `Order` value |
| `XLSX002` | A `[XlsxColumn]` member cannot be read or has no cell-value conversion |
| `XLSX003` | The row type is static, abstract, generic or not accessible |
| `XLSX005` | The row type has no readable columns |

If a type has no generated mapper, `InsertMappedTable` and `InsertMappedData` raise
`NotSupportedException`. Add `[XlsxTable]` to the type and make sure the generator runs.

### How values are written

| Member type | Cell value |
| --- | --- |
| `string` | Text |
| `bool` | Boolean |
| Integral and floating-point numbers, `decimal` | Number; `decimal` is stored as a double |
| `DateTime`, `DateOnly`, `DateTimeOffset` | Date/time serial value; the offset is dropped |
| `TimeSpan`, `TimeOnly` | Duration serial value |
| Enum | Text of the member name, for example `Shipped` |
| `Guid`, `char` | Text |
| `XlsxCellValue`, `XlsxError`, `Blank` | The value itself |
| `Nullable<T>` with `null` | Blank cell |

The reflection path handles `DateOnly` and `TimeOnly` differently: it calls `ToString()`,
so those values become text. Use the generated mapper when they must remain dates or
times. No text is parsed during import, so a string such as `"1,5"` stays text. The
`Format` code applies to data cells only, not to the header row, and does not change the
stored number. Culture does not affect stored values.

Empty input is accepted: `InsertMappedData(empty, addHeadings: true)` writes the header
row, and `InsertMappedTable(empty)` creates a header-only table with an empty data row.
