# Workbook I/O performance

The benchmarks use workbooks with one worksheet and five columns of numbers, repeated text,
dates, and formulas. Each load reads an XLSX file already in memory. Each save writes to a
new memory stream after loading the source workbook outside the timed operation. Package
validation and formula calculation are disabled during saving.

## Aid.Xlsx and ClosedXML on JIT

The standalone BenchmarkDotNet project compares Aid.Xlsx and ClosedXML 0.105.1 with both
libraries running on JIT. This command reproduces the settings of the Windows sample.

```sh
dotnet run --project tests/Aid.Xlsx.Benchmarks -c Release -- --filter '*' --job short --inProcess --artifacts artifacts/benchmarks
```

The Windows sample used BenchmarkDotNet 0.15.8 ShortRun with `--job short --inProcess` on
Windows 11, an AMD Ryzen 5 5600, and .NET 10.0.14. Times are means and allocations are
managed bytes per operation. The 2,000-row timings varied between iterations.

| Operation | Rows | ClosedXML time | Aid.Xlsx time | ClosedXML allocated | Aid.Xlsx allocated |
| --- | ---: | ---: | ---: | ---: | ---: |
| Load | 2,000 | 32.40 ms | 11.45 ms | 8.53 MB | 3.74 MB |
| Load | 20,000 | 218.35 ms | 131.88 ms | 79.91 MB | 34.19 MB |
| Save | 2,000 | 35.41 ms | 33.53 ms | 11.31 MB | 5.65 MB |
| Save | 20,000 | 425.23 ms | 162.45 ms | 105.92 MB | 47.77 MB |

## Aid.Xlsx on JIT and NativeAOT

Both builds run the same measurement program with the same XLSX inputs exported by the
BenchmarkDotNet project. The program reports median wall-clock time and mean managed bytes
allocated on the benchmark thread per operation after 20 warmup and 40 measured iterations
per case. The BenchmarkDotNet sample above uses a different measurement procedure and reports
mean times.

The sample used Windows 11 x64, an AMD Ryzen 5 5600, and .NET 10.0.14. We ran both
executables twice, reversing their order on the second run. Times show the range of
the two run medians. Allocations are averaged across both runs.

| Operation | Rows | JIT time | NativeAOT time | JIT allocated | NativeAOT allocated |
| --- | ---: | ---: | ---: | ---: | ---: |
| Load | 2,000 | 14.91–29.01 ms | 18.09–18.24 ms | 3.81 MiB | 3.74 MiB |
| Load | 20,000 | 128.94–140.34 ms | 205.98–206.93 ms | 34.18 MiB | 34.20 MiB |
| Save | 2,000 | 41.21–42.07 ms | 22.26–23.54 ms | 5.64 MiB | 5.27 MiB |
| Save | 20,000 | 179.44–180.06 ms | 239.71–241.51 ms | 47.76 MiB | 47.76 MiB |

JIT was faster on both 20,000-row operations. NativeAOT was faster when saving 2,000 rows.
The 2,000-row load result varied between runs. These measurements cover workbook I/O
after warmup and exclude process startup.

Run these commands from a Windows x64 Visual Studio Developer PowerShell with the C++
build tools installed.

```powershell
dotnet run --project tests/Aid.Xlsx.Benchmarks -c Release -- --export-input artifacts/native-benchmarks/input
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -o artifacts/native-benchmarks/jit
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -r win-x64 -p:PublishAot=true -p:NuGetLockFilePath=obj/packages.aot.lock.json -o artifacts/native-benchmarks/aot
.\artifacts\native-benchmarks\jit\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
.\artifacts\native-benchmarks\aot\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
```
