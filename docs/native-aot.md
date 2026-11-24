# NativeAOT and trimming

Aid.Xlsx 2.x is designed for trimmed and NativeAOT applications. The runtime
libraries carry `IsTrimmable` and `IsAotCompatible`, the trim and AOT analyzers are
enabled, and the workbook, OPC, XML, and formula paths avoid runtime reflection. The
remaining reflection-based entry points are annotated and have reflection-free
alternatives.

Use this guide when publishing with `PublishTrimmed` or `PublishAot`, when the build
reports `IL2xxx` warnings, or when the application imports row objects into a
workbook.

## Why the library supports trimming

The runtime libraries set these properties in `Directory.Build.props`:

| Property | Purpose |
| --- | --- |
| `IsTrimmable` | Marks the assembly as trim-compatible. |
| `IsAotCompatible` | Enables AOT analysis and supports NativeAOT use. |
| `EnableTrimAnalyzer` | Reports trim hazards at build time. |
| `EnableAotAnalyzer` | Reports dynamic-code and AOT hazards. |
| `EnableSingleFileAnalyzer` | Reports single-file packaging hazards. |
| `TrimmerSingleWarn` | Set to `false` so warnings identify individual call sites. |

The package does not take a runtime dependency on the Open XML SDK. Workbooks are
read and written by the library's own OPC/XML layer, and worksheet cells stream
through XML readers and writers. The source generator in the package emits table
mappers at compile time.

Applications still need to be analyzed. A library can be trim-compatible while an
application calls an annotated member. The analyzer warnings described below point
at those call sites.

## Use generated mappers for typed rows

The reflection-free path for typed records is `[XlsxTable]` plus
`InsertMappedTable` or `InsertMappedData`. The generator reads public instance fields
and readable properties and emits a mapper registered by a module initializer. No
reflection runs at startup or while writing rows.

```csharp
using System;
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Orders");

var orders = new[]
{
    new Order("Pears", 12.50m, new DateOnly(2026, 1, 2)),
    new Order("Apples", 8.00m, null)
};

sheet.Cell("A1").InsertMappedTable(orders);
sheet.Cell("F1").InsertMappedData(orders, addHeadings: true);
sheet.Cell("D1").FormulaA1 = "SUM(OrdersData[Amount])";

workbook.SaveAs("orders.xlsx", new SaveOptions
{
    EvaluateFormulasBeforeSaving = true
});

[XlsxTable("OrdersData")]
internal sealed record Order(
    [property: XlsxColumn(Header = "Item", Order = 1)] string Product,
    [property: XlsxColumn(Order = 2, Format = "0.00")] decimal Amount,
    DateOnly? DeliveredOn)
{
    [XlsxIgnore]
    public Uri? Link { get; init; }
}
```

`InsertMappedTable` creates an Excel table and uses the table name from
`[XlsxTable]` unless the optional `tableName` argument overrides it.
`InsertMappedData(data, addHeadings: true)` writes a plain range. Both accept an
`IEnumerable<T>` and write one row per item.

The generated mapper supports these member types, including their nullable forms:

| Kind | Types |
| --- | --- |
| Text and logic | `bool`, `string`, `char` |
| Integers | `sbyte`, `byte`, `short`, `ushort`, `int`, `uint`, `long`, `ulong` |
| Floating point | `float`, `double`, `decimal` |
| Date and time | `DateTime`, `DateTimeOffset`, `DateOnly`, `TimeOnly`, `TimeSpan` |
| Other values | `Guid`, `XlsxCellValue`, `XlsxError`, `Blank`, any enum except `XlsxError` |

Members must be readable public instance fields or properties. A private setter is
fine. Unsupported members produce compiler errors:

| Diagnostic | Meaning |
| --- | --- |
| `XLSX001` | Two members produce the same header or the same `Order` value. |
| `XLSX002` | A member is unreadable or has no supported type. |
| `XLSX003` | The row type is not a supported, accessible class or struct. |
| `XLSX005` | The row type has no readable columns. |

If a mapper is missing at run time, `InsertMappedTable` and `InsertMappedData` throw
`NotSupportedException` with the message `Add [XlsxTable] to 'Type' and enable the
Aid.Xlsx source generator.` Check that the source generator is referenced as an
analyzer and that the row type is annotated.

## The DataTable path

`DataTable` overloads are also reflection-free from the library's perspective and are
the right choice when rows come from a database, CSV import, or reporting layer.

```csharp
using System.Data;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

using var data = new DataTable("Stock");
data.Columns.Add("Item", typeof(string));
data.Columns.Add("Units", typeof(int));
data.Rows.Add("Pears", 12);
data.Rows.Add("Apples", 18);

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Stock");
var table = sheet.Cell("A1").InsertTable(data, "StockItems");
table.Theme = XlsxTableTheme.TableStyleMedium2;

workbook.SaveAs("stock.xlsx");
```

`InsertTable(DataTable, string)` creates a table. `InsertData(DataTable)` writes a
plain range, and `workbook.AddWorksheet(DataTable)` creates a sheet from a table. None
of these overloads is annotated with `RequiresUnreferencedCode`.

`DataTable` is part of the shared framework through `System.Data.Common`. It does not
add a package reference to a trimmed application.

## Reflection-based overloads and their warnings

The following members are annotated with `[RequiresUnreferencedCode]`. Their
attribute message names the supported alternatives.

| Overload | Replacement |
| --- | --- |
| `InsertData(IEnumerable)` | `InsertData(DataTable)`, `InsertMappedData`, or explicit cells. |
| `InsertTable<T>(IEnumerable<T>)` | `InsertTable(DataTable)` or `InsertMappedTable`. |
| `InsertTable(IEnumerable)` | `InsertTable(DataTable)` or `InsertMappedTable`. |
| `IXlsxTable.ReplaceData(IEnumerable)` | `ReplaceData(DataTable)` or write cells explicitly. |
| `IXlsxTable.AppendData(IEnumerable)` | `AppendData(DataTable)` or write cells explicitly. |

Calling one of these from code that is trimmed or compiled ahead of time produces a
trim warning such as `IL2026`. NativeAOT can also produce `IL3050` if the path reaches
dynamic code. The method still works on the full runtime, but the row type and its
members may be removed or altered after trimming.

Do not add `UnconditionalSuppressMessage` or `NoWarn` to silence these warnings.
Replace the call with a generated mapper or a `DataTable`, or keep the reflection
path out of the trimmed application entirely.

## Publish a trimmed or NativeAOT application

Prerequisites:

| Platform | Toolchain |
| --- | --- |
| Windows | .NET 11 SDK and Visual Studio Build Tools with the C++ workload and a Windows SDK. |
| Linux | .NET 11 SDK, `clang`, and the platform native build tools. |
| macOS | .NET 11 SDK and the Xcode command-line tools. |

From the application directory, publish for a runtime identifier:

```powershell
dotnet publish -c Release -r win-x64 -p:PublishAot=true
```

A trimmed, non-AOT publish is useful when a native toolchain is unavailable:

```powershell
dotnet publish -c Release -r win-x64 -p:PublishTrimmed=true
```

`PublishAot=true` implies trimming. `PublishTrimmed=true` without AOT still runs the
trim analyzer and removes unreachable code, but keeps the managed runtime.

### Optimization and globalization switches

```powershell
dotnet publish -c Release -r win-x64 -p:PublishAot=true -p:OptimizationPreference=Speed
```

`OptimizationPreference=Speed` spends more code size on faster generated code. The
repository measures it in the [performance guide](performance.md); it does not change
the API or the trimming behavior.

To remove ICU and culture data from a NativeAOT application, set
`InvariantGlobalization` in the project file:

```xml
<PropertyGroup>
  <InvariantGlobalization>true</InvariantGlobalization>
</PropertyGroup>
```

Caveats:

- The current culture becomes the invariant culture. `CultureInfo.GetCultureInfo` for
  a specific culture is unavailable, and `GetFormattedString(culture)` cannot render
  culture-specific number and date formats.
- Formula functions that parse culture-specific text, such as date and number parsing,
  operate with invariant rules. Workbooks whose formulas depend on a locale-specific
  literal can produce different results.
- If reports must show localized numbers or dates, keep ICU enabled and verify the
  output with the target runtime identifier instead of enabling this switch for size.

### Repository smoke publish

The checked-in smoke application exercises the whole library, including the
converter. From the repository root on Windows x64:

```powershell
dotnet publish tests/Aid.Xlsx.AotSmoke/Aid.Xlsx.AotSmoke.csproj `
  -c Release -r win-x64 -p:PublishAot=true `
  -p:NuGetLockFilePath=obj/packages.aot.lock.json `
  -o artifacts/aot/win-x64
./artifacts/aot/win-x64/Aid.Xlsx.AotSmoke.exe
```

The lock-file property is specific to this repository's locked restore. A normal
consumer application does not need it. The Linux and macOS equivalents use
`-r linux-x64` or `-r osx-arm64` and run the produced binary.

## Verified matrix

The repository publishes and executes these combinations:

| Pipeline | Runtime identifier | Build | Executed |
| --- | --- | --- | --- |
| `aot-smoke` in CI | `linux-x64` | `PublishTrimmed=true` | Yes |
| `aot-smoke` in CI | `linux-x64` | `PublishAot=true` | Yes |
| `aot-matrix` schedule | `linux-x64` | `PublishAot=true` | Yes |
| `aot-matrix` schedule | `win-x64` | `PublishAot=true` | Yes |
| `aot-matrix` schedule | `osx-arm64` | `PublishAot=true` | Yes |

No other runtime identifier is published by these workflows. In particular,
`osx-x64`, `linux-arm64`, and `win-arm64` are not covered by the repository matrix;
test your own target before shipping it.

The smoke application:

- creates and saves a workbook through a `MemoryStream`;
- inserts a generated-mapper table and a mapped data range;
- writes formulas, a number format, a date, a comment, a merged range, a validation
  rule, a hyperlink, and a conditional format;
- protects the worksheet and workbook with SHA-512 passwords, then reloads and
  unprotects them;
- converts an embedded BIFF8 `.xls` fixture and verifies formula and value results;
- rewrites the package to an artifact file when a path argument is supplied.

After publishing your own application, check that:

1. The publish output contains no `IL2xxx` warnings.
2. The native binary starts on a machine without the .NET runtime installed.
3. A workbook can be created, saved, reopened, and opened by Excel or LibreOffice.
4. Formula evaluation and number/date formatting match the JIT build for your data.
5. The binary size and startup time are acceptable for the chosen
   `OptimizationPreference`.

## Read trim and AOT diagnostics

Trim warnings are reported during publish. Each warning identifies the call site and
the annotated member that caused it.

| Warning | Meaning and action |
| --- | --- |
| `IL2026` | Call to a `RequiresUnreferencedCode` member; use a mapper or `DataTable`. |
| `IL3050` | Call to a `RequiresDynamicCode` member; avoid it in NativeAOT. |
| `IL2104` | Assembly summary warning; set `TrimmerSingleWarn=false`, then fix the sites. |

`TrimmerSingleWarn` is already `false` for the library projects. Set it in the
application when only the summary appears:

```xml
<PropertyGroup>
  <TrimmerSingleWarn>false</TrimmerSingleWarn>
</PropertyGroup>
```

Then publish again and search the output for `warning IL`:

```powershell
dotnet publish -c Release -r win-x64 -p:PublishAot=true 2>&1 | Select-String 'warning IL'
```

Read the warning from the bottom up. The first line names the method that needs
annotation or replacement; the call-site line names your code. A warning whose call
site is inside Aid.Xlsx is a library defect worth reporting; a warning whose call site
is in your code needs a change in your code.

For a warning-free publish, keep the trim and AOT analyzers enabled in the
application:

```xml
<PropertyGroup>
  <EnableTrimAnalyzer>true</EnableTrimAnalyzer>
  <EnableAotAnalyzer>true</EnableAotAnalyzer>
</PropertyGroup>
```

Treat warnings as errors only after the project is clean; turning them into errors
while known reflection paths remain makes the build fail without telling you which
alternative to use. The [troubleshooting guide](troubleshooting.md) covers the
`NotSupportedException` raised when a generated mapper is missing and the
[Cookbook](cookbook.md#nativeaot-safe-import) has a complete import sample.
