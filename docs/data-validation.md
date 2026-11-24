# Data validation

Data validation restricts what a user may type into a cell and can show an input prompt or an
error alert. It is enforced by the spreadsheet application while a user edits the sheet, not by
Aid.Xlsx while your code writes values. The rule is stored with the worksheet, so it survives
round trips and stays active when the file is opened.

```csharp
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Orders");

sheet.Cell("A1").Value = "Quantity";
sheet.Range("A2:A500").CreateDataValidation().WholeNumber.Between(1, 999);

workbook.SaveAs("orders.xlsx");
```

## Creating rules

`CreateDataValidation()` exists on `IXlsxCell`, on every range, and on an `IXlsxRanges`
collection. It always creates a new rule; to find an existing one, use `GetDataValidation()` or
the worksheet collection.

```csharp
using System;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Input");

// One rule over a whole column.
var quantity = sheet.Range("A2:A500").CreateDataValidation();
quantity.WholeNumber.Between(1, 999);

// A single cell needs a rule of its own.
sheet.Cell("C2").CreateDataValidation().List("Draft,Final,Archived");

// Reading rules back: TryGet covers the whole address, GetAllInRange only intersects it.
if (sheet.DataValidations.TryGet(sheet.Range("A2:A50").RangeAddress, out var found))
    found.ErrorTitle = "Quantity out of range";

foreach (var rule in sheet.DataValidations.GetAllInRange(sheet.Range("C2:C2").RangeAddress))
    Console.WriteLine(rule.AllowedValues);

workbook.SaveAs("input.xlsx");
```

- A worksheet keeps at most one rule per cell. Adding a rule to a range removes the overlapping
  area from every other rule and drops rules that lose all their cells, so ranges are split
  instead of nested.
- `IXlsxCell.GetDataValidation()` creates a rule when the cell has none; check
  `HasDataValidation` first when you only want to read. A range without a rule returns `null`
  from `GetDataValidation()`.
- `Ranges` lists the areas of one rule. `RemoveRange(range)` removes one area and returns
  whether it changed anything; `ClearRanges()` detaches all areas while keeping the settings.
- `Clear()` resets the rule to its defaults, which also means `AnyValue` and no messages, so a
  cleared rule has no effect.

### Removing rules

```csharp
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Cleanup");

sheet.Range("A2:A20").CreateDataValidation().List("Yes,No");

// Remove the rule only from the affected cells.
sheet.Range("A2:A20").Clear(XlsxClearOptions.DataValidation);

// Or delete every date rule on the sheet.
sheet.DataValidations.Delete(dv => dv.AllowedValues == XlsxAllowedValues.Date);

workbook.SaveAs("cleanup.xlsx");
```

`IXlsxDataValidations.Delete(predicate)` inspects the public rule state, so it can filter by
`AllowedValues`, `Operator`, `ErrorStyle`, `MinValue`, or `MaxValue`. Deleting a non-empty
range of cells through the grid API also deletes the rules that intersect the deleted area.

## Allowed types and operators

| `XlsxAllowedValues` | Meaning |
|---|---|
| `AnyValue` | No restriction, used for input messages only |
| `WholeNumber` | Integer value |
| `Decimal` | Floating point value |
| `List` | One value from a list |
| `Date` | Date serial value |
| `Time` | Time of day |
| `TextLength` | Length of the text in the cell |
| `Custom` | Formula that must evaluate to `TRUE` |

| `XlsxOperator` | Meaning |
|---|---|
| `Between`, `NotBetween` | Two values, minimum and maximum |
| `EqualTo`, `NotEqualTo` | One value |
| `GreaterThan`, `LessThan` | One value |
| `EqualOrGreaterThan`, `EqualOrLessThan` | One value |

`AllowedValues` and `Operator` can be set directly, but the typed criteria properties
(`WholeNumber`, `Decimal`, `Date`, `Time`, `TextLength`) set `AllowedValues` for you and
provide typed overloads for the comparison.

```csharp
using System;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Typed");

sheet.Range("A2:A100").CreateDataValidation().Decimal.Between(0.01, 9999.99);
sheet.Range("B2:B100").CreateDataValidation().WholeNumber.EqualOrGreaterThan(0);
sheet.Range("C2:C100").CreateDataValidation().TextLength.Between(2, 40);
sheet.Range("D2:D100").CreateDataValidation().Date.EqualOrGreaterThan(new DateTime(2026, 1, 1));
sheet.Range("E2:E100").CreateDataValidation()
    .Time.Between(TimeSpan.FromHours(8), TimeSpan.FromHours(18));

workbook.SaveAs("typed.xlsx");
```

- Reading a criteria property changes the rule: `dv.Date` assigns `AllowedValues = Date` even
  if you only wanted the criteria object, so order the calls so the type is set last if the
  rest of the configuration depends on it.
- `TextLength` accepts `int` values through the whole-number criteria base, because the length
  is an integer.
- Operator defaults to `Between`. For types that take a single value, any other operator can be
  used, but `Between` requires both a minimum and a maximum.
- Values are limited to 255 characters; longer strings throw `ArgumentOutOfRangeException` from
  the `MinValue`, `MaxValue`, and `Value` setters.

## Lists

A list rule accepts `List(string)`, which writes the source text as-is, or `List(range)`, which
writes an absolute reference to the range. A named range is passed as a string to the same
overload. Both forms have an overload that takes the in-cell dropdown flag.

```csharp
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Lists");

sheet.Cell("F1").Value = "Status";
sheet.Cell("F2").Value = "Draft";
sheet.Cell("F3").Value = "Final";

// Inline list: comma separated, limited to 255 characters in total.
sheet.Range("A2:A50").CreateDataValidation().List("Draft,Final,Archived");

// Range reference: the list source stays editable on the sheet.
var status = sheet.Range("B2:B50").CreateDataValidation();
status.List(sheet.Range("F2:F3"), inCellDropdown: true);

// Named range: define the name first, then pass it as the list source.
sheet.Range("F2:F3").AddToNamed("StatusValues");
sheet.Range("C2:C50").CreateDataValidation().List("StatusValues");

workbook.SaveAs("lists.xlsx");
```

- `List(range)` stores the absolute, sheet-qualified address of the range, so the list keeps
  working when rows or columns move inside that range. A reference to another sheet is written
  into the `x14` extension, which Excel reads fine and older consumers may ignore. A defined
  name stays in the main part as plain formula text.
- `InCellDropdown` is inverted in the file format (`showDropDown`); the property is the plain
  user-facing meaning, so `true` shows the arrow. The `List(range, inCellDropdown)` overload
  ignores its second argument in this version, so set `InCellDropdown` on the rule afterwards
  when the dropdown must be hidden.
- The dropdown arrow does not appear when the list source is empty, when the reference is
  invalid, or when the cell already holds a value that is not in the list.
- An inline list may not contain a comma in a value; use a range of cells for values with
  separators.

## Messages and alerts

Input messages appear when the cell is selected; error alerts appear when the user types a
value that fails the rule.

```csharp
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Price");

var price = sheet.Range("B2:B100").CreateDataValidation();
price.Decimal.Between(0.01, 9999.99);
price.ShowInputMessage = true;
price.InputTitle = "Unit price";
price.InputMessage = "Enter a price between 0.01 and 9999.99.";
price.ShowErrorMessage = true;
price.ErrorTitle = "Invalid price";
price.ErrorMessage = "Price must be a number between 0.01 and 9999.99.";
price.ErrorStyle = XlsxErrorStyle.Stop;
price.IgnoreBlanks = true;

workbook.SaveAs("price.xlsx");
```

| `XlsxErrorStyle` | Behavior in Excel |
|---|---|
| `Stop` | Rejects the value; Retry, Cancel, or Help |
| `Warning` | Offers Yes, No, or Cancel; Yes keeps the value |
| `Information` | Offers OK or Cancel |

- `ShowInputMessage` and `ShowErrorMessage` default to `true`, as do `IgnoreBlanks` and
  `InCellDropdown`; titles and messages default to an empty string.
- Title and message text are always stored; Excel displays them only when the corresponding
  `Show*` flag is `true`. An empty title with a non-empty message still shows the message,
  without a caption.
- A rule with `AnyValue` and no message text changes no behavior. When it is the only rule on
  the sheet, no `dataValidations` element is written at all.
- `IgnoreBlanks` allows an empty cell even when the rule would otherwise reject it. It does not
  allow an empty string typed into the cell when the rule requires a value.

## Criteria values and custom formulas

The `string` overloads of `XlsxValidationCriteria` write the text verbatim to `formula1` and
`formula2`, which is how values from other cells and formulas are expressed.

```csharp
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Criteria");

sheet.Cell("H1").Value = 100;

// Compare against another cell.
sheet.Range("A2:A20").CreateDataValidation()
    .Decimal.EqualOrLessThan(sheet.Cell("H1"));

// A formula rule: TRUE means the value is accepted.
sheet.Range("B2:B20").CreateDataValidation()
    .Custom("AND(ISNUMBER(B2),B2>0)");

workbook.SaveAs("criteria.xlsx");
```

- `Custom(formula)` sets `AllowedValues = Custom` and stores the expression as-is. Pass it
  without a leading `=`, write it for the first cell of the rule range, and let Excel shift the
  relative references to the other cells. See [Formulas and recalculation](formulas.md) for
  formula syntax and calculation.
- The cell overloads of the criteria methods write an absolute address of the compared cell,
  so the comparison stays fixed to that cell for the whole rule range.
- The date criteria convert `DateTime` to an OLE Automation date and write it with the invariant
  culture. Time criteria convert a `TimeSpan` to a fraction of a day.
- Formula-based criteria such as `INDEX` or `INDIRECT` work, but Excel recalculates the rule
  only when a dependent cell changes; the rule is not evaluated by Aid.Xlsx in any case.

## What validation does and does not do

- Programmatic writes are not checked. Assigning an out-of-range value with
  `cell.Value = ...` succeeds; Excel enforces the rule only when the user edits that cell.
  Validate in C# before you write when the data comes from an untrusted source; see
  [Cells and ranges](cells-and-ranges.md) for the typed `Get*` and `TryGetValue<T>` readers.
- Opening the file in Excel does not re-check existing values. Cells written outside the rule
  keep their values, and the rule applies to the next edit, paste, or fill.
- Circle-invalid markup is a UI feature of the application, not file content, so there is no
  Aid.Xlsx API to mark existing invalid values.
- A validated cell that is copied keeps the rule; copying a cell into an area that already has
  a rule replaces the overlapping part, following the one-rule-per-cell rule.
- `SaveOptions.ConsolidateDataValidationRanges` (default `true`) merges rules with identical
  settings and joins their ranges on save. It can reduce several added rules to one element
  with several `sqref` entries.

## Culture and formula pitfalls

- Strings passed to `MinValue`, `MaxValue`, `Value`, `Custom`, and `List` are written verbatim,
  so build them with the invariant culture. `"0.5"` is a valid criterion, `"0,5"` is a
  different, invalid one in the file.
- The typed `Decimal`, `WholeNumber`, and `Date` criteria already format their values
  invariantly, so prefer them over building strings by hand.
- `Time.Between` and the other `TimeSpan` overloads currently format the day fraction with the
  current culture, so a culture with a comma decimal separator produces an invalid criterion.
  Use the inherited `string` overload instead (`dv.Time.GreaterThan("0.5")`), or run the
  conversion under an invariant culture.
- A range reference in a criterion must be a valid address. Same-sheet references are written
  without the sheet name and cross-sheet references are written into the `x14` extension. A
  defined name is plain formula text and must already exist in the workbook.
- A custom formula must be valid A1-style syntax with invariant separators. Localized function
  names and `;` argument separators are not translated.
- Trimming whitespace from list values matters: `"Draft, Final"` allows `" Final"` with a
  leading space, while `"Draft,Final"` does not.
