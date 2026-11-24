# Workbook compatibility

Aid.Xlsx 1.x targets .NET 10 and reads, edits, and writes `.xlsx` and `.xlsm` packages.
Aid.Xlsx 2.x targets .NET 11 with the same API for the features described here. The
library uses its own OPC and SpreadsheetML implementation; the Open XML SDK is not a
runtime dependency.

The reader recognizes Transitional and Strict namespaces and relationships for supported
parts. Newly written workbooks use Transitional SpreadsheetML.

## Feature matrix

"Create" means the public API can add the feature to a workbook that did not have it.
"Edit" means an existing feature can be changed and saved. "Preserve" means the feature
survives a load and save without a public API to change it.

| Area | Read | Create | Edit | Notes |
| --- | --- | --- | --- | --- |
| Cell values, blanks, errors | Yes | Yes | Yes | Typed readers and `TryGetValue<T>`; see [Data types](data-types-and-conversions.md). |
| Formulas (A1, R1C1) | Yes | Yes | Yes | Cached results are read; recalculation is documented in [Formulas](formulas.md). |
| Styles: font, fill, border, alignment, number format | Yes | Yes | Yes | Styles are deduplicated on save. |
| Rich text runs | Yes | Yes | Yes | `CreateRichText` and `GetRichText`; see [Comments and rich text](comments-and-rich-text.md). |
| Comments (legacy notes) | Yes | Yes | Yes | One note per cell; threaded comments are not modeled. |
| Hyperlinks | Yes | Yes | Yes | External URI and internal references; see [Hyperlinks](hyperlinks-and-sparklines.md). |
| Defined names | Yes | Yes | Yes | Workbook and worksheet scope. |
| Tables | Yes | Yes | Yes | Themes, header and totals rows, auto-filters. |
| Auto-filters | Yes | Yes | Yes | Filter state is stored; filtering does not hide rows until an application applies it. |
| Conditional formatting | Yes | Yes | Yes | Cell rules, data bars, color scales, icon sets. |
| Data validation | Yes | Yes | Yes | Lists, numeric and date criteria, custom formulas. |
| Merged ranges | Yes | Yes | Yes | `Merge` and `Unmerge`; insertion can split partial merges. |
| Row and column layout | Yes | Yes | Yes | Widths, heights, hiding, outline levels. |
| Page setup and printing | Yes | Yes | Yes | Orientation, margins, print area, headers and footers. |
| Worksheet and workbook protection | Yes | Yes | Yes | Legacy password hash and SHA-512 for sheet protection. |
| Pivot tables and caches | Yes | Yes | Yes | `CreatePivotTable` from a range; caches can be refreshed on open. |
| Pictures | Yes | Yes | Yes | Add from path or stream, move, scale, delete. |
| Charts | Preserve | No | No | Charts are kept as package data. `IXlsxChart` has no public factory and no public worksheet collection, so charts cannot be created or edited through the API. |
| VBA project in `.xlsm` | Preserve | No | No | Treated as opaque data and kept byte-for-byte on save. |
| Unsupported package parts | Preserve | No | No | Kept as opaque data when the containing part is not rewritten. |
| Encryption (password to open) | No | No | No | The library does not decrypt or encrypt workbook packages. |
| Legacy `.xls` (BIFF8) | Convert | No | No | `XlsConverter` reads BIFF8 and writes `.xlsx`; see below. |

## Preservation rules

- Parts outside the workbook model can survive a save as opaque data. Round-trip tests
  cover VBA projects, disconnected custom parts, relationships, and unsupported DrawingML
  shapes.
- Editing a modeled part can rewrite its XML and affect markup that the object model does
  not represent. Saving can change ZIP entry order and timestamps, XML prefixes and
  formatting, relationship targets, and calculation chain entries.
- A chart inside a worksheet does not block cell edits. Deleting or rewriting a drawing
  that contains an unsupported shape is not supported.
- `SaveOptions.ValidatePackage` checks internal links and supported XML roots after the
  package has been written. The test suite additionally validates generated fixtures with
  the Open XML SDK schemas; the runtime package checks the OPC structure with its own
  implementation.

## Legacy `.xls` conversion

`XlsConverter` accepts Excel 97–2003 BIFF8 workbooks and produces `.xlsx`.

| Content | Conversion |
| --- | --- |
| Values, dimensions, sheet visibility, merged ranges | Copied. |
| Number formats, fonts, palette colors, alignment, borders, fills, cell protection flags | Copied. |
| Formulas with supported operators, same-sheet and same-workbook references, common functions such as `SUM`, `IF`, `ROUND`, and `VLOOKUP` | Translated; the cached result is preserved. |
| Shared formulas, array formulas, references spanning multiple sheets or workbooks | Written as static values. |
| Named styles, rich text runs, conditional formatting, images, charts, macros, encrypted files | Outside the converter scope. |

Operation is all-or-nothing per file: convert to a temporary result first, then replace
the destination. An existing destination is replaced only after a successful conversion.
See [Workbook files and streams](workbook-files.md#convert-a-legacy-xls-workbook).

## Verification in this repository

- The public API, including domain namespaces, is compared with the checked-in baseline
  `tests/Aid.Xlsx.ApiCompat/Baselines/PublicApi.txt`. Run the compatibility project with
  `--print-api` to inspect the declarations after an intentional API change.
- Ten corpus workbooks are loaded, saved, and reopened; the results are compared with
  checked-in reference snapshots, and the saved packages are validated with the SDK
  schemas.
- CI publishes the corpus for manual checks: `dotnet run --project tests/Aid.Xlsx.ApiCompat
  -c Release -- --save-corpus artifacts/workbook-corpus`. Excel 2021 and LibreOffice Calc
  opened and resaved the corpus, including pivot tables, charts, pictures, conditional
  formats, and validation rules.
- Round-trip tests compare VBA project bytes and unknown part payloads before and after a
  save.

## Limits to plan around

- Workbook size is bounded by Excel limits: 1,048,576 rows and 16,384 columns per
  worksheet, 32,767 characters per cell, and 255 characters for a header or footer line.
- A workbook must contain at least one visible worksheet before it can be saved.
- Charts, macros, and other opaque parts are preserved but not editable.
- The calculation engine covers a subset of Excel functions and does not evaluate
  external workbook references. Check [Formulas](formulas.md) before relying on
  values computed outside Excel.
- Protection controls editing in spreadsheet applications; it is not encryption.
  See [Protection](protection.md).

## See also

- [Migration from ClosedXML](migration-from-closedxml.md) for renames and replacements.
- [Troubleshooting](troubleshooting.md) for the exceptions these limits produce.
- [Workbook files and streams](workbook-files.md) for load and save options.
