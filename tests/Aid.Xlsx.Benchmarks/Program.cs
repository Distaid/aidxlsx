using Aid.Xlsx.Benchmarks;
using BenchmarkDotNet.Running;
using ClosedXML.Excel;

if (args is ["--export-input", var directory])
{
    Directory.CreateDirectory(directory);
    File.WriteAllBytes(Path.Combine(directory, "rows-2000.xlsx"), WorkbookIoBenchmarks.CreateInput(2_000));
    File.WriteAllBytes(Path.Combine(directory, "rows-20000.xlsx"), WorkbookIoBenchmarks.CreateInput(20_000));
    return;
}

if (args is ["--measure-closedxml", var inputDirectory])
{
    var saveOptions = new SaveOptions { ValidatePackage = false, EvaluateFormulasBeforeSaving = false };
    WorkbookIoMeasurement.Run(inputDirectory, "ClosedXML",
        static source => new XLWorkbook(source),
        static (workbook, rows) => workbook.Worksheet(1).Cell(rows, 1).GetDouble(),
        (workbook, output) => workbook.SaveAs(output, saveOptions));
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(WorkbookIoBenchmarks).Assembly).Run(args);
