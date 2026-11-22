using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ClosedXML.Excel;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Benchmarks;

[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public class WorkbookIoBenchmarks
{
    private static readonly Aid.Xlsx.Model.Workbooks.SaveOptions AidSaveOptions = new() { ValidatePackage = false };
    private static readonly ClosedXML.Excel.SaveOptions BaselineSaveOptions = new() { ValidatePackage = false };
    private XlsxWorkbook? _workbook;
    private XLWorkbook? _baseline;
    private MemoryStream? _saveInput;
    private byte[] _input = null!;

    [Params(2_000, 20_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        using var seed = new XLWorkbook();
        var sheet = seed.AddWorksheet("Transactions");
        for (var row = 1; row <= Rows; row++)
        {
            sheet.Cell(row, 1).Value = row;
            sheet.Cell(row, 2).Value = $"Item {row % 97}";
            sheet.Cell(row, 3).Value = row / 10d;
            sheet.Cell(row, 4).Value = new DateTime(2026, 1, 1).AddDays(row % 365);
            sheet.Cell(row, 5).FormulaA1 = $"C{row}*2";
        }

        sheet.Column(3).Style.NumberFormat.Format = "0.00";
        sheet.Column(4).Style.DateFormat.Format = "yyyy-mm-dd";
        using var input = new MemoryStream();
        seed.SaveAs(input, BaselineSaveOptions);
        _input = input.ToArray();
    }

    [IterationSetup(Target = nameof(ClosedXmlSave))]
    public void PrepareClosedXmlSave()
    {
        _saveInput = OpenInput();
        _baseline = new XLWorkbook(_saveInput);
    }

    [IterationSetup(Target = nameof(AidXlsxSave))]
    public void PrepareAidXlsxSave()
    {
        _saveInput = OpenInput();
        _workbook = new XlsxWorkbook(_saveInput);
    }

    [IterationCleanup(Targets = [nameof(ClosedXmlSave), nameof(AidXlsxSave)])]
    public void CleanupSave()
    {
        _workbook?.Dispose();
        _baseline?.Dispose();
        _saveInput?.Dispose();
        _workbook = null;
        _baseline = null;
        _saveInput = null;
    }

    [Benchmark(Baseline = true), BenchmarkCategory("Load")]
    public double ClosedXmlLoad()
    {
        using var input = OpenInput();
        using var workbook = new XLWorkbook(input);
        return workbook.Worksheet(1).Cell(Rows, 1).GetDouble();
    }

    [Benchmark, BenchmarkCategory("Load")]
    public double AidXlsxLoad()
    {
        using var input = OpenInput();
        using var workbook = new XlsxWorkbook(input);
        return workbook.Worksheet(1).Cell(Rows, 1).GetDouble();
    }

    [Benchmark(Baseline = true), BenchmarkCategory("Save")]
    public long ClosedXmlSave()
    {
        using var output = new MemoryStream();
        _baseline!.SaveAs(output, BaselineSaveOptions);
        return output.Length;
    }

    [Benchmark, BenchmarkCategory("Save")]
    public long AidXlsxSave()
    {
        using var output = new MemoryStream();
        _workbook!.SaveAs(output, AidSaveOptions);
        return output.Length;
    }

    private MemoryStream OpenInput() => new(_input, writable: false);
}
