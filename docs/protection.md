# Protection and file security

Aid.Xlsx exposes three different mechanisms that are easy to confuse: worksheet and
workbook protection, cell-level `Locked` and `Hidden` flags, and a read-only
recommendation in file sharing. None of them encrypts the workbook. This guide shows what
each one does, how passwords are stored, and where encryption is not available.

## Worksheet protection

Use worksheet protection to keep users from changing formulas, structure or formatting
while leaving selected cells editable. Unlock the input cells first, then protect the
sheet with the elements you want to allow.

```csharp
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Input");

sheet.Cell("A1").Value = "Name";
sheet.Cell("B1").Value = "Amount";
sheet.Cell("A2").Value = "Pears";
sheet.Cell("B2").Value = 12.5;

// Unlock the cells users may edit before enabling protection.
sheet.Range("A2:B100").Style.Protection.Locked = false;

sheet.Protect(
    password: "change-me",
    algorithm: XlsxProtectionAlgorithm.Algorithm.SHA512,
    allowedElements: XlsxSheetProtectionElements.SelectLockedCells
        | XlsxSheetProtectionElements.SelectUnlockedCells);

workbook.SaveAs("input.xlsx");
```

`Protect()` without a password protects the sheet but lets anyone turn protection off.
`Protect(password)`, `Protect(allowedElements)`, `Protect(algorithm, allowedElements)` and
`Protect(password, algorithm, allowedElements)` cover the other combinations. The call
returns `IXlsxSheetProtection`, which exposes `AllowElement`, `DisallowElement`,
`AllowEverything`, `AllowNone` and `AllowedElements` for later adjustments.

`sheet.IsProtected` reports whether protection is on, and `sheet.IsPasswordProtected`
reports whether it requires a password. `sheet.Unprotect()` works only when no password is
set; `sheet.Unprotect("change-me")` removes password-protected protection.

Protect throws `InvalidOperationException` when the sheet is already protected. Unprotect
throws `InvalidOperationException` when a password is required but none is supplied, and
`ArgumentException` when the password is wrong.

### Allowed elements

The flags name what a user may still do while the sheet is protected. Anything not allowed
becomes prohibited. The password overload without an `allowedElements` argument defaults to
`SelectEverything`, so only selection remains possible.

| `XlsxSheetProtectionElements` member | Allows |
| --- | --- |
| `SelectLockedCells`, `SelectUnlockedCells` | Selecting locked or unlocked cells |
| `FormatCells`, `FormatColumns`, `FormatRows` | Changing cell, column or row formatting |
| `InsertColumns`, `InsertRows`, `InsertHyperlinks` | Inserting columns, rows or hyperlinks |
| `DeleteColumns`, `DeleteRows` | Deleting columns or rows |
| `Sort`, `AutoFilter`, `PivotTables` | Sorting, filtering, or changing pivot reports |
| `EditObjects`, `EditScenarios` | Editing drawing objects or scenarios |
| `None`, `Everything` | Nothing, or every supported operation |
| `DeleteEverything`, `FormatEverything`, `InsertEverything`, `SelectEverything` | Combined groups |

`AllowEverything()` sets every flag but leaves the sheet marked as protected, so the file
still contains a `sheetProtection` element even though it no longer restricts anything.
`AllowNone()` clears every flag and prohibits all listed operations.

### Password hashing

`XlsxProtectionAlgorithm.DefaultProtectionAlgorithm` is `SimpleHash`, the legacy 16-bit
scheme Excel has used for years. `XlsxProtectionAlgorithm.Algorithm` has two values:

| Algorithm | How the password is stored | Limitation |
| --- | --- | --- |
| `SimpleHash` | A 16-bit hash in the `password` attribute | Short, unsalted and fast to crack |
| `SHA512` | Salted SHA-512 hash with 100,000 iterations | Slower; still just a password check |

`SHA512` writes `algorithmName`, `hashValue`, `saltValue` and `spinCount` into the
`sheetProtection` element; `SimpleHash` writes only `password`. The legacy scheme converts
the password with an ASCII encoder, so characters outside ASCII are not represented
faithfully. Neither algorithm encrypts the sheet contents, and neither protects data at
rest.

## Workbook protection

Workbook protection restricts the workbook structure: inserting, deleting, renaming,
moving or hiding sheets. It does not restrict editing cells and does not lock the file.

```csharp
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("template.xlsx");

workbook.Protect(
    password: "structure",
    algorithm: XlsxProtectionAlgorithm.Algorithm.SHA512,
    allowedElements: XlsxWorkbookProtectionElements.None);

// Protection is advisory for Excel; the library still edits the model.
workbook.Worksheet(1).Cell("A1").Value = "Updated";
workbook.SaveAs("locked.xlsx");
```

| `XlsxWorkbookProtectionElements` member | Allows |
| --- | --- |
| `Structure` | Changing the sheet structure |
| `Windows` | Changing the window layout |
| `None` | Locking both structure and windows |
| `Everything` | Allowing both, which leaves nothing locked |

The password overload without `allowedElements` allows `Windows` and therefore locks the
structure. To lock both, pass `XlsxWorkbookProtectionElements.None`.

`workbook.LockStructure` and `workbook.LockWindows` read and write the same flags. Their
setters throw `InvalidOperationException` until `Protect` has been called. Use
`workbook.Unprotect()` or `workbook.Unprotect(password)` to turn protection off.

Workbook protection is not a security boundary. It does not stop another application from
editing the file, and Aid.Xlsx itself does not check protection when you assign values.

### Read protection state in an existing file

Opening a workbook never fails because a worksheet or the structure is protected; the
library only reads the protection settings. Inspect them before editing a template.

```csharp
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Workbooks;

using IXlsxWorkbook workbook = new XlsxWorkbook("protected.xlsx");
var sheet = workbook.Worksheet("Input");

Console.WriteLine(sheet.IsProtected);
Console.WriteLine(sheet.IsPasswordProtected);
Console.WriteLine(sheet.Protection.AllowedElements);

Console.WriteLine(workbook.IsProtected);
Console.WriteLine(workbook.Protection.AllowedElements);
```

`IsPasswordProtected` is `true` only when a non-empty password hash is present. A sheet
protected without a password reports `IsProtected = true` and
`IsPasswordProtected = false`. You can unprotect and protect again to change a password:

```csharp
sheet.Unprotect("old-password");
sheet.Protect("new-password", XlsxProtectionAlgorithm.Algorithm.SHA512);
```

## Cell-level protection

Every cell has two protection flags. `Locked` defaults to `true` and `Hidden` defaults to
`false`. The flags only take effect when the worksheet is protected.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var report = workbook.Worksheets.Add("Report");

report.Cell("A1").Value = "Quantity";
report.Cell("B2").Value = 10;

// Users may edit these cells while the sheet is protected.
report.Range("B2:B100").Style.Protection.Locked = false;

// Hide the formula from the formula bar; the calculated value stays visible.
report.Cell("C2").FormulaA1 = "B2*2";
report.Cell("C2").Style.Protection.Hidden = true;

report.Protect("secret");
workbook.SaveAs("report-protected.xlsx");
```

`Locked = false` makes a cell editable, but the element flags still apply: formatting,
sorting or filtering can remain prohibited for the whole sheet. `Hidden = true` hides the
formula in the formula bar while the cell value and the grid display remain visible. It
does not hide the cell and does not protect its content. Without sheet protection, both
flags have no visible effect.

See [Formatting and worksheet features](formatting-and-worksheets.md) for styling cells
through `Style`, and [Cells and ranges](cells-and-ranges.md) for selecting the ranges you
want to unlock.

## Read-only recommendation and file sharing

`IXlsxFileSharing` adds the recommendation Excel shows when a workbook is opened. It is a
prompt, not a restriction: the user can choose to edit and save normally.

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("review.xlsx");
workbook.FileSharing.ReadOnlyRecommended = true;
workbook.FileSharing.UserName = "Review Desk";
workbook.SaveAs("review-shared.xlsx");
```

| `IXlsxFileSharing` member | Effect |
| --- | --- |
| `ReadOnlyRecommended` | Writes `readOnlyRecommended="1"` so Excel suggests read-only mode |
| `UserName` | Writes the `userName` attribute shown with the recommendation |

The public interface does not expose a write-reservation password, algorithm name, hash
value, salt or spin count. Those parts of the OOXML `fileSharing` element are neither read
nor written, so a password that prevents saving changes is not available in Aid.Xlsx.

`workbook.IsPasswordProtected` refers to workbook structure protection described above, not
to a password required to open the file.

## Encrypted workbooks are not supported

Aid.Xlsx reads and writes ordinary OPC ZIP packages. It has no API that accepts a password
for opening or saving a workbook: neither `XlsxWorkbook` constructors nor `LoadOptions` and
`SaveOptions` expose one. A workbook saved by Aid.Xlsx is a normal `.xlsx` package whose XML
parts can be read by any ZIP tool.

A password-to-open workbook is an OLE compound file that contains an encrypted package, not
a ZIP package. Opening one throws `InvalidDataException` from the packaging layer:

```csharp
using System.IO;
using Aid.Xlsx.Model.Workbooks;

try
{
    using var workbook = new XlsxWorkbook("password-protected.xlsx");
}
catch (InvalidDataException exception)
{
    Console.WriteLine($"This file is not a plain OPC package: {exception.Message}");
}
```

The library does not decrypt the package, does not prompt for a password and does not write
encrypted output. To protect confidential workbooks, encrypt the file outside Aid.Xlsx, for
example with an encrypted archive or the storage layer that holds it, and keep access
control outside the document.

## What protection does not do

- It does not encrypt the package or any XML part; values remain readable by ZIP and XML
  tools.
- It does not stop Aid.Xlsx or another library from editing the workbook; the flags are
  interpreted by the spreadsheet application.
- It does not prevent copying, printing, screenshots or exporting data through another
  program.
- It does not hide cell values. `Hidden` only suppresses the formula in the formula bar.
- It does not authenticate a user. There is no identity or permission model behind the
  password hash.

## Security guidance

- Treat worksheet and workbook protection as protection against accidental edits, not as
  access control. The content stays readable in the package.
- Prefer `SHA512` when you need an Excel-compatible protection password, but do not reuse a
  sensitive password: the hash is still stored in the file and can be attacked offline.
- Do not store secrets, credentials or personal data in a protected sheet. Remove or redact
  them before the workbook leaves your application.
- Validate submitted values on the server. A protected input sheet does not guarantee that
  a client sent valid or expected data.
- Use `Hidden` only to reduce clutter in the formula bar, never to conceal a formula that
  must remain secret.
- For confidential distribution, combine external encryption with access control and data
  minimization; document protection alone does not provide confidentiality.
