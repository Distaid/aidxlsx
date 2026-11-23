# Workbook files and streams

## Edit a workbook or save a copy

Use `Save()` to write to the current source. Use `SaveAs(...)` to write to another
path or stream, or to save a new workbook for the first time. When editing a template
or an existing report, keep the source by saving to a different path.

Use any existing `.xlsx` with at least one worksheet as `source.xlsx`. You can also
copy the workbook created in the [getting started guide](getting-started.md).

```csharp
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("source.xlsx");
var sheet = workbook.Worksheet(1);
sheet.Cell("A1").Value = "Revised report";

workbook.SaveAs("revised.xlsx", new SaveOptions { ValidatePackage = true });
```

`ValidatePackage` checks the workbook relationships and the structure of supported
XML parts. See [Workbook compatibility](compatibility.md) for the scope of these checks.

Keep the `.xlsm` extension when editing a macro-enabled workbook and writing a copy
that should retain its VBA project. For example, open `source.xlsm` and save to
`revised.xlsm`. Aid.Xlsx treats the VBA project as opaque data during workbook edits.

## Work with streams

Supply a readable, seekable stream positioned at the beginning of the workbook.
Keep it open while working with the workbook, especially when saving a copy. The
caller owns the streams. Use a fresh seekable output stream and rewind it before
reading the result.

```csharp
using System.IO;
using Aid.Xlsx.Model.Workbooks;

using var input = File.OpenRead("source.xlsx");
using var output = new MemoryStream();

using (var workbook = new XlsxWorkbook(input))
{
    workbook.Worksheet(1).Cell("A1").Value = "Written to a stream";
    workbook.SaveAs(output);
}

output.Position = 0;
using var copy = File.Create("revised.xlsx");
output.CopyTo(copy);
```

Use a writable source stream if you plan to save back to it with `Save()`. An
application that writes to a non-seekable response can save into a `MemoryStream`
and copy the result to the response.

## Convert legacy Excel files

`XlsConverter` converts Excel 97–2003 BIFF8 workbooks to `.xlsx`. The one-argument
overload changes the extension and uses the same directory. It replaces an existing
destination after the workbook has been converted successfully.

```csharp
using Aid.Xlsx.Conversion;

XlsConverter.Convert("legacy.xls");
XlsConverter.Convert("legacy.xls", "converted.xlsx");
```

For `XlsConverter.Convert(sourceStream, destinationStream)`, use a readable,
seekable source and a separate writable destination. Both remain open. The supported
formatting and formula subset is described in [Workbook compatibility](compatibility.md).
