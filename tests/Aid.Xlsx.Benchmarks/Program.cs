using Aid.Xlsx.Benchmarks;
using BenchmarkDotNet.Running;

if (args is ["--export-input", var directory])
{
    Directory.CreateDirectory(directory);
    File.WriteAllBytes(Path.Combine(directory, "rows-2000.xlsx"), WorkbookIoBenchmarks.CreateInput(2_000));
    File.WriteAllBytes(Path.Combine(directory, "rows-20000.xlsx"), WorkbookIoBenchmarks.CreateInput(20_000));
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(WorkbookIoBenchmarks).Assembly).Run(args);
