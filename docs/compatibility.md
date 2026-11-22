# Workbook compatibility

Aid.Xlsx reads and edits `.xlsx` and `.xlsm` packages. Its workbook model covers cell
values and formulas, styles, rich text, tables, names, conditional formatting, validation,
notes, hyperlinks, images, print layout, protection, pivot tables, and pivot caches.
The reader recognizes Transitional and Strict namespaces and relationships for supported
parts. Newly written workbooks use Transitional SpreadsheetML.

Parts outside the workbook model can survive a save as opaque data. Round-trip tests
cover VBA projects, disconnected custom parts, relationships, and unsupported DrawingML
shapes. You can edit worksheet cells in a workbook containing a chart. Chart creation and
editing have no public API. Editing a modeled part can rewrite its XML and affect markup
that the object model does not represent.

Saving can change ZIP entry order and timestamps, XML prefixes and formatting,
relationship targets, and calculation chain entries. `SaveOptions.ValidatePackage`
checks internal links and supported XML roots. The test suite uses Open XML SDK schema
validation for generated fixtures. The runtime package checks the OPC structure using
its own implementation.

`XlsConverter` accepts Excel 97–2003 BIFF8 workbooks. It copies values, dimensions,
visibility, and merged ranges. Cell formatting includes number formats,
fonts, indexed or custom palette colors, alignment, borders, fills, and cell protection flags.
It translates formulas with supported operators, cell and range references on individual sheets in the
same workbook, and common worksheet functions such as `SUM`, `IF`, `ROUND`, and `VLOOKUP`, preserving
the cached result. Shared formulas, array formulas, references spanning multiple sheets or pointing to
another workbook, and other unsupported formulas become static values.
Named styles, rich text runs, conditional formatting, images, charts, macros, and encrypted
files are outside the converter's supported scope.

## Spreadsheet application checks

Export the workbook corpus used by CI into a directory for checking in spreadsheet applications.

```sh
dotnet run --project tests/Aid.Xlsx.ApiCompat -c Release -- --save-corpus artifacts/workbook-corpus
```

Excel 2021 for Windows (16.0.14332.20303) and LibreOffice Calc for Windows (26.2.7.3)
opened and resaved the ten corpus files and the NativeAOT smoke workbook. In Excel,
the checks also verified the pivot table, embedded chart, picture, conditional format
and validation rule. VBA execution was disabled during the Excel check. The library's
tests separately compare the VBA project bytes before and after saving with Aid.Xlsx.

## Public API comparison

CI compares workbook API declarations against the published ClosedXML 0.105.1 package.
The comparison maps the namespace and type renames, checks signatures and declared
members, and records the existing differences in `tests/Aid.Xlsx.ApiCompat/AcceptedDifferences.txt`.

- Package and DrawingML helpers that exposed SDK types are absent from the workbook API.
- The imported workbook model differs on a few low-level types, including style cache
  types and arithmetic helpers. It also differs on some obsolete validation APIs,
  range collections, sparkline state and picture disposal.
- Aid.Xlsx adds compiled table mapping, the BIFF8 converter, load options and automatic
  colors. It also records nullable annotations on a few existing members.
