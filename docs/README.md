# Aid.Xlsx documentation

Aid.Xlsx reads, creates, and edits Excel workbooks (`.xlsx`) and macro-enabled workbooks
(`.xlsm`) from .NET without the Open XML SDK at runtime. This index groups the guides by
task and by workbook area.

Every public type, member, and namespace mentioned in these guides is taken from the
checked-in API baseline in `tests/Aid.Xlsx.ApiCompat/Baselines/PublicApi.txt`. If an
example and the baseline disagree, the baseline is authoritative.

## Start here

| Guide | What it covers |
| --- | --- |
| [Getting started](getting-started.md) | Install the package, create, reopen, and update a workbook. |
| [Cookbook](cookbook.md) | Short, complete recipes for common workbook tasks. |
| [Migration from ClosedXML](migration-from-closedxml.md) | Rename types, map namespaces, and find replacements for missing APIs. |
| [Troubleshooting](troubleshooting.md) | Read exceptions, diagnose failed saves and loads, and file a useful issue. |

## Workbook and file handling

| Guide | What it covers |
| --- | --- |
| [Workbook files and streams](workbook-files.md) | Load and save options, streams, templates, `.xlsm`, workbook properties, defined names, and legacy `.xls` conversion. |
| [Compatibility](compatibility.md) | Feature matrix, preservation rules, conversion limits, validation scope, and how CI verifies workbooks. |
| [Performance](performance.md) | Measured load and save results, plus guidance for large workbooks and many styles. |
| [NativeAOT and trimming](native-aot.md) | Compile-time row mapping, trim warnings, and publishing trimmed or NativeAOT applications. |

## Cells and ranges

| Guide | What it covers |
| --- | --- |
| [Cells and ranges](cells-and-ranges.md) | Address cells, traverse ranges, merge, insert and delete rows and columns, copy and move, and sort. |
| [Data types and conversions](data-types-and-conversions.md) | Values, blanks, errors, dates, precision, `Get*` and `TryGetValue<T>`, and culture-sensitive formatting. |
| [Formulas and recalculation](formulas.md) | A1 and R1C1 formulas, cached values, recalculation, and formula errors. |

## Formatting and worksheet features

| Guide | What it covers |
| --- | --- |
| [Formatting and worksheet features](formatting-and-worksheets.md) | Fonts, fills, borders, alignment, number formats, styles, row and column sizing, freeze panes, page setup, and printing. |
| [Conditional formatting](conditional-formatting.md) | Cell rules, top and bottom rules, data bars, color scales, icon sets, and formulas. |
| [Data validation](data-validation.md) | Lists, numeric and date rules, custom formulas, input messages, and error alerts. |

## Content and objects

| Guide | What it covers |
| --- | --- |
| [Comments and rich text](comments-and-rich-text.md) | Notes on cells, comments, and run-level rich text formatting. |
| [Hyperlinks and sparklines](hyperlinks-and-sparklines.md) | External and internal links, tooltips, and sparkline groups. |
| [Images and drawings](images-and-drawings.md) | Pictures, anchors, placement, scaling, z-order, and preservation of unsupported drawings. |

## Tables and analysis

| Guide | What it covers |
| --- | --- |
| [Tables and data import](tables-and-data-import.md) | Create and style tables, auto-filters, `DataTable` import, and compile-time mapped records. |
| [Pivot tables](pivot-tables.md) | Read, edit, and preserve pivot tables, caches, fields, and styles. |
| [Protection](protection.md) | Worksheet and workbook protection, passwords, cell locking, and read-only recommendations. |

## Conventions in these guides

- Examples are console programs that target `net10.0` (Aid.Xlsx 1.x). For `net11.0` and
  the 2.x package line, the API is the same unless a guide says otherwise.
- `using` declarations dispose workbooks and streams deterministically. Keep source
  streams alive while the workbook uses them.
- A workbook needs at least one worksheet before it can be saved.
- Row and column indexes start at `1`, so `Cell(1, 1)` is `A1`.
- Number and date format codes use Excel's invariant syntax. Culture affects display
  and parsing, not the stored format code.
- Code samples use the `Xlsx*` and `IXlsx*` names; the `IXlsx*` interfaces are the
  return types of the factory methods such as `Cell`, `Range`, and `AddWorksheet`.

## See also

- [Repository README](../README.md) for building, testing, and packaging.
- [API compatibility baseline](../tests/Aid.Xlsx.ApiCompat/Baselines/PublicApi.txt) for
  the complete public surface.
- [Workbook compatibility](compatibility.md) for what is preserved on a round trip.
