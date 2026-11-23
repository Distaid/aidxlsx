# Getting started

## Install Aid.Xlsx

Create a console application with the .NET 10 SDK.

```sh
dotnet new console --framework net10.0 --name WorkbookExample
cd WorkbookExample
```

Once `1.0.0` is available on NuGet.org, add it from your application directory.

```sh
dotnet add package Aid.Xlsx --version 1.0.0
```

To try the current source, pack the library from the repository root.

```sh
dotnet pack src/Aid.Xlsx/Aid.Xlsx.csproj -c Release -o artifacts/packages
```

From your application directory, use the version in the generated `.nupkg` filename
and the absolute path to the local package directory.

```sh
dotnet add package Aid.Xlsx --version 1.0.0 --no-restore
dotnet restore --source "<absolute-path-to-artifacts/packages>" --source https://api.nuget.org/v3/index.json
```

## Create a workbook

Replace `Program.cs` with this example. A workbook needs at least one worksheet
before it can be saved. The `using` declaration disposes it when the program exits.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Sales");

sheet.Cell("A1").Value = "Item";
sheet.Cell("B1").Value = "Units";
sheet.Cell(2, 1).Value = "Pears";
sheet.Cell(2, 2).Value = 12;

workbook.SaveAs("sales.xlsx");
Console.WriteLine("Created sales.xlsx.");
```

Run `dotnet run --no-restore` from the application directory. The workbook appears
in the current working directory.

## Reopen and edit the workbook

Replace `Program.cs` with this example and run it again from the same directory. It
prints `12` and writes a copy with the quantity changed to `15`.

```csharp
using System;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook("sales.xlsx");
var sheet = workbook.Worksheet("Sales");

Console.WriteLine(sheet.Cell("B2").GetDouble());
sheet.Cell("B2").Value = 15;
workbook.SaveAs("sales-updated.xlsx");
```

To keep building the workbook, see [Cells and ranges](cells-and-ranges.md) and
[Formatting and worksheet features](formatting-and-worksheets.md).
