# Aid.Xlsx.CodeGen

This command-line tool generates forward-only SpreadsheetML readers from the pinned ECMA-376
Transitional XSD. The generated partial classes handle XML structure and attribute conversion.
Handwritten partial classes own the hooks that validate inputs and populate the workbook model.

| Directory | Responsibility |
| --- | --- |
| [Schema](Schema/) | Fetches and verifies the pinned XSD. `Parsing/` builds the schema representation in `Model/`. |
| [SpreadsheetReaders](SpreadsheetReaders/) | Selects the schema elements and C# mappings for each workbook reader |
| [Emission](Emission/) | Emits parsing methods, conversion calls, and partial hook declarations |
| [Infrastructure](Infrastructure/) | Helper types used by the model and emitter |

## Regenerate the readers

From the repository root run:

```sh
dotnet run --project src/Aid.Xlsx.CodeGen -- all
```

The first run downloads and verifies the schema. Subsequent runs reuse the verified copy under
`artifacts/codegen/`. The command updates the four `.g.cs` files next to their handwritten
counterparts in the main project. [SpreadsheetML](../Aid.Xlsx/SpreadsheetML/)

| Command | Output under `SpreadsheetML/` |
| --- | --- |
| `styles` | `Styles/StylesReader.g.cs` |
| `cache-records` | `Pivots/PivotCacheRecordsReader.g.cs` |
| `partial-rst` | `Strings/RstReader.g.cs` |
| `partial-sheet-data` | `Worksheets/SheetDataReader.g.cs` |

For debugging a single recipe, pass `<command> <schema.xsd> <output.cs>` after `--`. To change a
reader, edit the recipe or its handwritten counterpart, regenerate, and review the diff.

## Configure a recipe

`SchemeTypeMap` maps XSD simple types to C# attribute conversions. The required template reads an
attribute or reports a missing or invalid value. The optional template returns `null` for an
absent attribute. Use `{0}` as a placeholder when the template includes the attribute name.

`AddComplexTypeMapping` specifies the C# return type for a parsed element so a parent can consume
the value. `AddParseMethod` selects the elements for code generation. After consuming an element,
the generated method passes its attributes and child values to a handwritten hook. Implement the
hook when the generated method needs to return a mapped value.
