using Aid.Xlsx.Benchmarks;
using Aid.Xlsx.Model.Workbooks;

if (args is not [var inputDirectory])
{
    Console.Error.WriteLine("Usage: Aid.Xlsx.NativeBenchmarks <input-directory>");
    return 1;
}

var saveOptions = new SaveOptions { ValidatePackage = false, EvaluateFormulasBeforeSaving = false };
WorkbookIoMeasurement.Run(inputDirectory, "Aid.Xlsx",
    static source => new XlsxWorkbook(source),
    static (workbook, rows) => workbook.Worksheet(1).Cell(rows, 1).GetDouble(),
    (workbook, output) => workbook.SaveAs(output, saveOptions));
return 0;
