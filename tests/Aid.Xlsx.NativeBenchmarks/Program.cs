using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Aid.Xlsx.Model.Workbooks;

if (args is not [var inputDirectory])
{
    Console.Error.WriteLine("Usage: Aid.Xlsx.NativeBenchmarks <input-directory>");
    return 1;
}

const int warmupCount = 20;
const int sampleCount = 40;
var saveOptions = new SaveOptions { ValidatePackage = false, EvaluateFormulasBeforeSaving = false };
Console.WriteLine($"{(RuntimeFeature.IsDynamicCodeSupported ? "JIT" : "NativeAOT")} | {RuntimeInformation.FrameworkDescription} | {RuntimeInformation.ProcessArchitecture}");
Console.WriteLine("Operation,Rows,MedianMilliseconds,MeanAllocatedBytes");

foreach (var rows in new[] { 2_000, 20_000 })
{
    var input = File.ReadAllBytes(Path.Combine(inputDirectory, FormattableString.Invariant($"rows-{rows}.xlsx")));
    var loadTimes = new double[sampleCount];
    var saveTimes = new double[sampleCount];
    long loadBytes = 0;
    long saveBytes = 0;

    for (var index = -warmupCount; index < sampleCount; index++)
    {
        var allocated = GC.GetAllocatedBytesForCurrentThread();
        var started = Stopwatch.GetTimestamp();
        double lastValue;
        using (var source = new MemoryStream(input, writable: false))
        using (var workbook = new XlsxWorkbook(source))
            lastValue = workbook.Worksheet(1).Cell(rows, 1).GetDouble();
        var elapsed = Stopwatch.GetElapsedTime(started);
        var bytes = GC.GetAllocatedBytesForCurrentThread() - allocated;
        if (lastValue != rows)
            throw new InvalidDataException("Unexpected value in the last row of the input workbook.");
        if (index >= 0)
        {
            loadTimes[index] = elapsed.TotalMilliseconds;
            loadBytes += bytes;
        }
    }

    for (var index = -warmupCount; index < sampleCount; index++)
    {
        using var source = new MemoryStream(input, writable: false);
        using var workbook = new XlsxWorkbook(source);
        var allocated = GC.GetAllocatedBytesForCurrentThread();
        var started = Stopwatch.GetTimestamp();
        using (var output = new MemoryStream())
        {
            workbook.SaveAs(output, saveOptions);
            if (output.Length == 0)
                throw new InvalidDataException("The save operation produced an empty workbook.");
        }
        var elapsed = Stopwatch.GetElapsedTime(started);
        var bytes = GC.GetAllocatedBytesForCurrentThread() - allocated;
        if (index >= 0)
        {
            saveTimes[index] = elapsed.TotalMilliseconds;
            saveBytes += bytes;
        }
    }

    PrintResult("Load", rows, loadTimes, loadBytes);
    PrintResult("Save", rows, saveTimes, saveBytes);
}

return 0;

static void PrintResult(string operation, int rows, double[] times, long allocatedBytes)
{
    Array.Sort(times);
    var median = (times[(times.Length - 1) / 2] + times[times.Length / 2]) / 2;
    var meanAllocated = (double)allocatedBytes / times.Length;
    Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"{operation},{rows},{median:F2},{meanAllocated:F0}"));
}
