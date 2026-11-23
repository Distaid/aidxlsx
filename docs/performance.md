# Workbook I/O performance

The general I/O benchmarks use workbooks with one worksheet and five columns of numbers,
repeated text, dates, and formulas. Each load reads an XLSX file already in memory. Each
save writes to a new memory stream after loading the source workbook outside the timed
operation. Package validation and formula calculation are disabled during saving. The
general I/O results include the style ID index introduced in commit `10b8cb1`.

## Aid.Xlsx and ClosedXML on JIT

The standalone BenchmarkDotNet project compares Aid.Xlsx and ClosedXML 0.105.1 with both
libraries running on JIT. This command reproduces the settings of the Windows sample.

```sh
dotnet run --project tests/Aid.Xlsx.Benchmarks -c Release -- --filter '*' --job short --inProcess --artifacts artifacts/benchmarks
```

The Windows sample used BenchmarkDotNet 0.15.8 ShortRun with `--job short --inProcess` on
Windows 11, an AMD Ryzen 5 5600, and .NET 10.0.14. Times are means and allocations are
managed bytes per operation. The short save iterations for 2,000 rows had high timing
variation.

| Operation | Rows | ClosedXML time | Aid.Xlsx time | ClosedXML allocated | Aid.Xlsx allocated |
| --- | ---: | ---: | ---: | ---: | ---: |
| Load | 2,000 | 30.35 ms | 11.03 ms | 8.53 MB | 3.74 MB |
| Load | 20,000 | 229.83 ms | 148.09 ms | 79.92 MB | 34.19 MB |
| Save | 2,000 | 38.86 ms | 49.88 ms | 11.31 MB | 5.65 MB |
| Save | 20,000 | 443.44 ms | 164.85 ms | 105.92 MB | 47.77 MB |

## ClosedXML on JIT and Aid.Xlsx on JIT and NativeAOT

ClosedXML 0.105.1 on JIT and all Aid.Xlsx builds use the same measurement code and XLSX
inputs exported by the BenchmarkDotNet project. The program reports median wall-clock time
and mean managed bytes allocated on the benchmark thread per operation after 20 warmup
and 40 measured iterations per case. The BenchmarkDotNet sample above uses a different
measurement procedure and reports mean times.

The sample used Windows 11 x64, an AMD Ryzen 5 5600, and .NET 10.0.14. We ran all four
builds twice and reversed their order for the second run. Times show the range of the
two run medians. Allocations are averaged across both runs.

### Wall-clock time per operation

| Operation | Rows | ClosedXML JIT | Aid.Xlsx JIT | Aid.Xlsx NativeAOT | Aid.Xlsx NativeAOT `Speed` |
| --- | ---: | ---: | ---: | ---: | ---: |
| Load | 2,000 | 27.26–62.79 ms | 14.09–14.92 ms | 17.76–17.94 ms | 15.73–15.79 ms |
| Load | 20,000 | 210.25–216.62 ms | 127.81–128.16 ms | 207.51–214.62 ms | 190.82–191.62 ms |
| Save | 2,000 | 40.00–57.30 ms | 42.12–42.28 ms | 21.82–21.85 ms | 20.56–21.34 ms |
| Save | 20,000 | 401.58–411.49 ms | 179.42–183.20 ms | 235.04–239.00 ms | 226.44–229.19 ms |

### Managed allocations per operation

| Operation | Rows | ClosedXML JIT | Aid.Xlsx JIT | Aid.Xlsx NativeAOT | Aid.Xlsx NativeAOT `Speed` |
| --- | ---: | ---: | ---: | ---: | ---: |
| Load | 2,000 | 8.74 MiB | 3.78 MiB | 3.74 MiB | 3.74 MiB |
| Load | 20,000 | 79.91 MiB | 34.18 MiB | 34.20 MiB | 34.20 MiB |
| Save | 2,000 | 11.20 MiB | 5.64 MiB | 5.27 MiB | 5.27 MiB |
| Save | 20,000 | 105.90 MiB | 47.76 MiB | 47.76 MiB | 47.76 MiB |

On 20,000 rows, default NativeAOT load times overlapped with ClosedXML on JIT. Both
NativeAOT builds saved that workbook faster than ClosedXML. `OptimizationPreference=Speed`
reduced the NativeAOT medians in all four cases.
[NativeAOT optimization settings](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/optimizing)

The measurements cover workbook I/O after warmup and exclude process startup.
NativeAOT avoids JIT compilation at process startup. A warmed-up JIT can use runtime
profiles to optimize frequently executed code paths, which may help explain the results.
The results describe this Windows configuration. [Tiered compilation and dynamic PGO](https://learn.microsoft.com/en-us/dotnet/core/runtime-config/compilation)

Run these commands from a Windows x64 Visual Studio Developer PowerShell with the C++
build tools installed. Repeat the last four commands in reverse order for the second run.

```powershell
dotnet run --project tests/Aid.Xlsx.Benchmarks -c Release -- --export-input artifacts/native-benchmarks/input
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -o artifacts/native-benchmarks/jit
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -r win-x64 -p:PublishAot=true -p:NuGetLockFilePath=obj/packages.aot.lock.json -o artifacts/native-benchmarks/aot
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -r win-x64 -p:PublishAot=true -p:OptimizationPreference=Speed -p:NuGetLockFilePath=obj/packages.aot.lock.json -o artifacts/native-benchmarks/aot-speed
dotnet run --project tests/Aid.Xlsx.Benchmarks -c Release --no-build -- --measure-closedxml artifacts/native-benchmarks/input
.\artifacts\native-benchmarks\jit\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
.\artifacts\native-benchmarks\aot\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
.\artifacts\native-benchmarks\aot-speed\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
```

## Saving workbooks with many styles

Commit `10b8cb1` added an index for saved style IDs. We compared it with its parent commit
using Release builds on JIT under .NET 10.0.14 on Windows 11 x64. The synthetic workbook
contains 10,000 rows with one numeric cell per row. Each cell has a distinct fill color,
with 512 font colors and 200 custom number formats in rotation.

Each run used 8 warmup saves and 25 measured saves. It loaded a fresh copy of the workbook
before each save and wrote to a new memory stream. Times are medians and allocations are
mean managed bytes on the benchmark thread per save. Package validation and formula
calculation were disabled. We ran each build twice and reversed their order for the second
pair of runs.

| Build | Median save times | Mean allocated per save |
| --- | ---: | ---: |
| Before the style index | 218.40–230.83 ms | 1.31 GB |
| With the style index | 58.06–65.09 ms | 21.17–21.27 MB |

For comparison, we ran the JIT benchmark above on the standard 20,000-row workbook. The
two save medians were 190.88 and 237.26 ms before the change, and 189.88 and 189.97 ms
with the index. Both builds allocated about 50.08 MB per save. The spread between the
two baseline runs makes the timing improvement on this input inconclusive.
