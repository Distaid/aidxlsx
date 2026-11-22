using Aid.Xlsx.Benchmarks;
using BenchmarkDotNet.Running;

BenchmarkSwitcher.FromAssembly(typeof(WorkbookIoBenchmarks).Assembly).Run(args);
