# NativeAOT and trimming

Use the `Aid.Xlsx` package from a .NET 10 application. The package includes the table
mapper generator. Annotate each concrete row type with `[XlsxTable]` and call
`InsertMappedData` or `InsertMappedTable`. Public instance fields and readable properties
become columns. `[XlsxColumn]` controls heading, order, and number format. `[XlsxIgnore]`
omits a member.

```csharp
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.DataImport.Mapping;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Orders");
sheet.Cell("A1").InsertMappedTable(new[] { new Order("Pears", 12.50m) });
workbook.SaveAs("orders.xlsx");

[XlsxTable("Orders")]
internal record Order(string Item, [property: XlsxColumn(Format = "0.00")] decimal Total);
```

Legacy `InsertData` and `InsertTable` overloads for arbitrary objects can inspect their
members with reflection. Publish with the trim and AOT analyzers enabled and address
their warnings at calls to those overloads. Collections of explicit rows, such as arrays
of cell values, can also avoid object reflection.

From the repository root, publish and run the smoke application for the current OS and
architecture. This example uses Windows x64 and the Visual Studio C++ toolchain.

```powershell
dotnet publish tests/Aid.Xlsx.AotSmoke -c Release -r win-x64 -p:PublishAot=true -p:NuGetLockFilePath=obj/packages.aot.lock.json -o artifacts/aot/win-x64
./artifacts/aot/win-x64/Aid.Xlsx.AotSmoke.exe
```

Main CI publishes and runs both fully trimmed and NativeAOT builds on Linux x64. The
scheduled NativeAOT matrix runs on Linux x64, Windows x64, and macOS Arm64.
