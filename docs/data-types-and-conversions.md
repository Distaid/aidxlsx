# Data types and conversions

A cell value is a `XlsxCellValue`: a small struct that holds exactly one of seven data
types. `IXlsxCell.DataType` tells you which one is present, and `IXlsxCell.Value` exposes
the value. Reads are either strict (`GetText`, `GetDouble`) or converting
(`TryGetValue<T>`, `GetValue<T>`).

## Cell value model

Use `DataType` when the content is mixed and you must branch before reading. Use the
`Is*` properties on `XlsxCellValue` when you already hold a value.

```csharp
using System;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Types");
sheet.Cell("A1").Value = Blank.Value;
sheet.Cell("A2").Value = true;
sheet.Cell("A3").Value = 1234.5;
sheet.Cell("A4").Value = "Text";
sheet.Cell("A5").Value = XlsxError.NoValueAvailable;
sheet.Cell("A6").Value = new DateTime(2026, 1, 15);
sheet.Cell("A7").Value = new TimeSpan(1, 30, 0);

foreach (string address in new[] { "A1", "A2", "A3", "A4", "A5", "A6", "A7" })
{
    IXlsxCell cell = sheet.Cell(address);
    Console.WriteLine($"{address}: {cell.DataType} = {cell.GetString()}");
}
```

| `XlsxDataType` | Value | Typical source |
| --- | --- | --- |
| `Blank` | No value | Unset cell, `Blank.Value`, `null` in `FromObject`. |
| `Boolean` | `true` or `false` | `bool` value. |
| `Number` | IEEE 754 double | Any numeric value or `decimal`. |
| `Text` | String, up to 32,767 characters | `string`, rich text, converted objects. |
| `Error` | `XlsxError` | Formula errors and explicit error values. |
| `DateTime` | Date and time | `DateTime`, date-formatted numbers. |
| `TimeSpan` | Elapsed time | `TimeSpan`, time-formatted numbers. |

Rich text is not a separate `XlsxDataType`: a cell with formatting runs is `Text`.
`HasRichText`, `CreateRichText()`, and `GetRichText()` work on the text; `GetText()`
returns the plain string. A formula cell reports the type of its cached or evaluated
result, not a separate formula type.

## Assign values

Use `Value` for a literal and `FormulaA1`/`FormulaR1C1` for a formula. `Value` always
writes a value: assigning text that begins with `=` does not create a formula, and
assigning any value to a formula cell removes the formula.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Assign");
sheet.Cell("A1").Value = 42;              // Number
sheet.Cell("A2").Value = "00123";         // Text, keeps leading zeros
sheet.Cell("A3").Value = "'00456";        // Text "00456"; apostrophe is removed
sheet.Cell("A4").Value = 0.1m;            // Decimal converted to double
sheet.Cell("A5").Value = (int?)null;      // Blank
sheet.Cell("A6").Value = "=1+1";          // Text, not a formula

sheet.Cell("B2").FormulaA1 = "A1*2";
sheet.Cell("B2").Value = 100;                   // The formula is replaced
Console.WriteLine(sheet.Cell("B2").HasFormula); // False
```

`bool`, `DateTime`, `TimeSpan`, `Blank`, and `XlsxError` keep their type. All numeric
types, including `decimal`, become `Number` and are converted to double. A `string`
becomes `Text`; a null string becomes `Blank`. Nullable numeric, `DateTime`, and
`TimeSpan` values become `Blank` when null. `FromObject` maps supported values and falls
back to `Convert.ToString` for anything else.

- `XlsxWorkbook.CellSetValueBehavior` selects `Smart` (0) or `Simple` (1). The `Value`
  property is typed `XlsxCellValue`, so text is never re-parsed into a number or date on
  assignment. A numeric-looking string stays text; assign the typed value instead.
- A leading apostrophe is stripped only when the value is a string, so `"'00456"` stores
  `00456` as text. The apostrophe is not an escape that survives in the cell.
- `SetValue(XlsxCellValue)` does the same as `Value =` and returns the cell. The setter
  also applies a matching date or time number format and updates table headers.
- A write through an inferior cell of a merged range is silently ignored.
- `double.NaN` and infinity throw `ArgumentException`; text longer than 32,767 characters
  and error values outside the enum throw `ArgumentOutOfRangeException`.

## Read typed values

Use a strict getter when you control the file and know the type. Use a converting getter
when the file comes from a user or another application.

```csharp
using System;
using System.Globalization;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Read");
sheet.Cell("A1").Value = 1234.5;
sheet.Cell("A2").Value = new DateTime(2026, 1, 15);
sheet.Cell("A3").Value = "00123";
sheet.Cell("A4").Value = true;
IXlsxCell number = sheet.Cell("A1");
number.Style.NumberFormat.Format = "#,##0.00";

Console.WriteLine(number.GetDouble());              // 1234.5
Console.WriteLine(number.GetFormattedString());     // number format applied
Console.WriteLine(number.GetFormattedString(CultureInfo.InvariantCulture));
Console.WriteLine(sheet.Cell("A2").GetDateTime().Year); // 2026
Console.WriteLine(sheet.Cell("A3").GetText());      // 00123
Console.WriteLine(sheet.Cell("A4").GetBoolean());   // True
```

| Member | Result | Throws |
| --- | --- | --- |
| `GetText()` | Text exactly as stored. | `InvalidCastException` for any other type. |
| `GetString()` | Current-culture display text; blank becomes `""`. | Nothing for normal values. |
| `GetFormattedString()` | Formatted display text. | Falls back to `CachedValue`. |
| `GetDouble()` | The number. | `InvalidCastException` for text, date, or time. |
| `GetBoolean()` | The logical value. | `InvalidCastException` for any other type. |
| `GetDateTime()` | The date and time. | `InvalidCastException` for a plain number. |
| `GetTimeSpan()` | The elapsed time. | `InvalidCastException` for any other type. |
| `GetError()` | The error code. | `InvalidCastException` for a non-error cell. |
| `GetValue<T>()` | Converted value. | `InvalidCastException` when conversion fails. |
| `TryGetValue<T>(out T)` | `true` when conversion succeeded. | Rarely; failures return `false`. |

There is no `GetInt32()`, and `GetValue<object>()` does not produce a boxed value. Use
`GetValue<int>()`, `GetString()`, or `TryGetValue<int>(out int value)`.
`GetFormattedString(CultureInfo)` takes an explicit culture; the parameterless overload
uses the current culture. `GetDouble()` is strict: on a date cell it throws, even though
dates are stored as a serial number internally. Use `GetValue<double>()` to obtain it.

## TryGetValue and conversion rules

Use `TryGetValue<T>` for import code and any path where a single bad cell must not stop
the program. It catches failures from formula evaluation and from conversion.

```csharp
using System;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Convert");
sheet.Cell("A1").Value = "42";
sheet.Cell("A2").Value = 42.75;
sheet.Cell("A3").Value = "";
sheet.Cell("A4").Value = XlsxError.DivisionByZero;
if (sheet.Cell("A1").TryGetValue<int>(out int fromText))
    Console.WriteLine(fromText); // 42, parsed with the current culture

if (!sheet.Cell("A2").TryGetValue<int>(out _))
    Console.WriteLine("42.75 is not an integer.");
if (sheet.Cell("A3").TryGetValue<int?>(out int? missing))
    Console.WriteLine(missing.HasValue); // False, empty text becomes a null value
if (sheet.Cell("A4").TryGetValue<XlsxError>(out XlsxError error))
    Console.WriteLine(error); // DivisionByZero
```

| Target `T` | Conversion |
| --- | --- |
| `string` | Any value becomes current-culture text; `_xHHHH_` escapes are decoded. |
| Nullable value type | Blank and empty text become `null`; otherwise the inner type is used. |
| `DateTime` | `DateTime` directly; number or `TimeSpan` when it is a valid serial date. |
| `TimeSpan` | `TimeSpan`; number as serial time; text parsed with the current culture. |
| `bool` | Boolean; number (`0` is false); text `TRUE`/`FALSE`, case-insensitive. |
| `XlsxError` | Only error values. |
| Enum | Enum name from the display text, case-sensitive. |
| `float`, `double`, `decimal` | Number, boolean, numeric text, or a date serial number. |
| Integer types | Must be integral and inside the target type's range; otherwise `false`. |

`TryGetValue` uses `CultureInfo.CurrentCulture` for text parsing and display. A formula
that throws while being evaluated makes it return `false` instead of propagating.

## Blank, empty text, and errors

Use `Blank.Value` to clear a value and to compare against an unset cell. Inspect
`DataType` to distinguish an empty string from a blank cell, and use `GetError()` only
after you know the cell holds an error.

```csharp
using System;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Empty");
var neverSet = sheet.Cell("A1");
var emptyText = sheet.Cell("A2");
var error = sheet.Cell("A3");

emptyText.Value = "";
error.Value = XlsxError.CellReference;

Console.WriteLine(neverSet.DataType);          // Blank
Console.WriteLine(neverSet.IsEmpty());         // True
Console.WriteLine(emptyText.DataType);         // Text
Console.WriteLine(emptyText.GetText().Length); // 0
Console.WriteLine(emptyText.IsEmpty());        // True
Console.WriteLine(error.IsEmpty());            // False
Console.WriteLine(error.GetError());           // CellReference
Console.WriteLine(error.GetString());          // #REF!
```

| `XlsxError` | Numeric | Display text |
| --- | --- | --- |
| `NullValue` | 0 | `#NULL!` |
| `DivisionByZero` | 1 | `#DIV/0!` |
| `IncompatibleValue` | 2 | `#VALUE!` |
| `CellReference` | 3 | `#REF!` |
| `NameNotRecognized` | 4 | `#NAME?` |
| `NumberInvalid` | 5 | `#NUM!` |
| `NoValueAvailable` | 6 | `#N/A` |

- `Blank.Value` is a singleton; `XlsxCellValue` exposes `IsBlank`, `GetBlank()`, and
  `TryConvert(out Blank value)`, which succeeds for blank and empty text.
- An empty string is `Text`, not `Blank`, but `IsEmpty()` treats it as empty. A formula
  that returns an empty string is not empty because `HasFormula` is true.
- `GetString()` and `GetFormattedString()` render an error as its `#` code. `GetText()`
  and the numeric getters throw `InvalidCastException` on an error cell.
- Excel error codes the library does not model, such as `#SPILL!` or `#CALC!`, leave the
  cached value blank. Write only the seven values above.
- A cast such as `(XlsxError)42` throws `ArgumentOutOfRangeException` on assignment
  because only values 0 through 6 exist. `cell.Value = 3` is the number three, not
  `#REF!`; use `XlsxError.CellReference` for the error.

## Dates and times

`DateTime` and `TimeSpan` values are stored as Excel serial numbers. The in-memory model
always uses the 1900 date system, and the reader and writer translate 1904 files for you;
use `Use1904DateSystem` only to preserve the source workbook's date-system setting.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Dates");
sheet.Cell("A1").Value = new DateTime(1900, 2, 28);
sheet.Cell("A2").Value = new DateTime(1900, 3, 1);
sheet.Cell("A3").Value = new DateTime(2026, 1, 15, 13, 45, 30);
sheet.Cell("A4").Value = new TimeSpan(2, 15, 0, 0);
Console.WriteLine(sheet.Cell("A1").GetValue<double>()); // 59
Console.WriteLine(sheet.Cell("A2").GetValue<double>()); // 61
Console.WriteLine(sheet.Cell("A3").GetDateTime().Hour); // 13
Console.WriteLine(sheet.Cell("A4").GetTimeSpan().TotalHours); // 63

workbook.Use1904DateSystem = true; // mark 1904; values stay 1900-based
```

- Excel reserves serial 60 for a nonexistent 1900-02-29. Reading a cell with that serial
  as a date throws `ArgumentException`; `DateTime` cannot represent the value.
- Excel serial dates round to the nearest millisecond when converted back to `DateTime`.
  `TimeSpan` uses ticks, but writing it to a double can still lose sub-millisecond digits.
- `DateTime.Kind` and an offset are not stored. The serial number is a local wall-clock
  value. Normalize to UTC in the application and keep the time zone separately.
- `DateTimeOffset` is not a cell type. Assign its `DateTime` part explicitly; passing it
  to `XlsxCellValue.FromObject` falls back to text.
- A `TimeSpan` is a fraction of a day; `2.15:00` is stored as `2.625`. Values beyond 24
  hours are valid and display according to the cell's number format.
- `Use1904DateSystem` is written to the workbook, and the reader adds 1,462 days on load;
  the value your code reads is already the correct calendar date.

## Numbers, precision, and culture

Use `double` for every number. Excel and this library store numbers as binary floating
point, so exact decimal fractions and integers beyond 2^53 are not preserved. Use an
explicit `CultureInfo` when the display or parsing culture must not follow the process.

```csharp
using System;
using System.Globalization;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Numbers");
var german = CultureInfo.GetCultureInfo("de-DE");
sheet.Cell("A1").Value = 9007199254740993L; // rounds to the nearest double
sheet.Cell("A2").Value = 0.1m;              // decimal becomes double
sheet.Cell("A3").Value = "00042";           // leading zeros require text
sheet.Cell("A4").Value = "1,5";             // numeric-looking text stays text
sheet.Cell("A4").Style.NumberFormat.Format = "#,##0.00";

CultureInfo.CurrentCulture = german;
if (sheet.Cell("A4").TryGetValue<double>(out double parsed))
    Console.WriteLine(parsed); // 1.5
```

- Integers up to 9,007,199,254,740,991 are exact. Larger integers lose low-order digits,
  and a `decimal` is rounded to the nearest double on assignment.
- Excel displays at most 15 significant digits, but the file stores the full double.
  `XlsxWorkbook.FullPrecision` is a calculation flag written to the workbook; it does not
  change how this library stores values.
- Store identifiers, postal codes, and other leading-zero values as text. Text that looks
  numeric is not included in numeric aggregation unless a function explicitly coerces it.
- Number formats change display only. `GetDouble()` returns the stored number regardless
  of the format; `GetFormattedString(culture)` applies the format and the culture.
- Text is limited to 32,767 characters per cell; a longer string throws
  `ArgumentOutOfRangeException` at assignment.
- `TryGetValue<T>` parses text with the current culture: a decimal comma is a number in
  `de-DE` and a separator in `en-US`. The file itself stores values invariantly.
- The calculation engine also uses the current culture for text-to-number coercion; see
  [Formulas](formulas.md).

## Round-tripping

Use this section to decide which values must be re-derived after a save and reopen. Values
survive, but their source representation and some edge digits do not.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using (var workbook = new XlsxWorkbook())
{
    var sheet = workbook.Worksheets.Add("RoundTrip");
    sheet.Cell("A1").Value = 0.1m;
    sheet.Cell("A2").Value = new DateTime(2026, 1, 15, 13, 45, 30, 123);
    sheet.Cell("A3").Value = 12345678901234567890m;
    sheet.Cell("A4").Value = "00123";
    sheet.Cell("A5").Value = XlsxError.NameNotRecognized;
    workbook.SaveAs("roundtrip.xlsx");
}
using var reopened = new XlsxWorkbook("roundtrip.xlsx");
var read = reopened.Worksheet("RoundTrip");
Console.WriteLine(read.Cell("A1").GetDouble());
Console.WriteLine(read.Cell("A2").GetDateTime().Millisecond); // 123
Console.WriteLine(read.Cell("A3").GetDouble());               // precision was reduced
Console.WriteLine(read.Cell("A4").GetText());                 // 00123
Console.WriteLine(read.Cell("A5").GetError());                // NameNotRecognized
```

| Item | Preserved |
| --- | --- |
| Blank, boolean, number, text, error | Yes; the data type is stored in the file. |
| Date and time | To the millisecond; `Kind` and offset are not stored. |
| Time span | As a serial fraction of a day, subject to double precision. |
| Leading zeros | When stored as text. |
| Rich text runs | Yes; `HasRichText` and `GetRichText()` survive a round-trip. |
| Number format and style | Yes; see [Formatting](formatting-and-worksheets.md). |
| Formula text | Yes, but cached results depend on save options. |
| `decimal` exactness, integers above 2^53 | No; both are stored as double. |
| Unknown error codes | No; unmodeled codes read back as blank. |

- Formula cells written without evaluation lose their cached result. Reading such a cell
  triggers evaluation if the formula can be evaluated; `CachedValue` remains blank.

## Exceptions and limits

Use this table to decide which failure mode to handle. Conversion failures are ordinary
`InvalidCastException`; assignment failures are argument exceptions.

```csharp
using System;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Limits");

try
{
    sheet.Cell("A1").Value = new string('x', 32768);
}
catch (ArgumentOutOfRangeException exception)
{
    Console.WriteLine(exception.Message);
}
```

| Operation | Exception |
| --- | --- |
| Text longer than 32,767 characters | `ArgumentOutOfRangeException`. |
| `NaN`, positive infinity, or negative infinity | `ArgumentException`. |
| Error value outside `NullValue..NoValueAvailable` | `ArgumentOutOfRangeException`. |
| Strict getter on the wrong data type | `InvalidCastException`. |
| `GetValue<T>` when conversion fails | `InvalidCastException`. |
| Very large double converted to `decimal` | `OverflowException`. |
| `TryGetValue<T>` on any mismatch | No exception; it returns `false`. |

- `GetFormattedString()` falls back to `CachedValue` when evaluating a formula throws, so
  a malformed or unsupported formula does not break display formatting.
- `XlsxCellValue.FromObject` maps `null`, `Blank`, `bool`, `string`, `XlsxError`,
  `DateTime`, `TimeSpan`, and the numeric types. Anything else becomes text through
  `Convert.ToString`, which is culture-sensitive.

For addressing and traversal see [Cells and ranges](cells-and-ranges.md); for save
behavior see [Workbook compatibility](compatibility.md).
