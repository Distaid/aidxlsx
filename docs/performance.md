# Workbook I/O performance

The general I/O benchmarks use workbooks with one worksheet and five columns of numbers,
repeated text, dates, and formulas. Each load reads an XLSX file already in memory. Each
save writes to a new memory stream after loading the source workbook outside the timed
operation. Package validation and formula calculation are disabled during saving.

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

### NativeAOT optimized for speed

A follow-up run on the same Windows machine used the same inputs and ran the default
NativeAOT build, NativeAOT with `OptimizationPreference=Speed`, and JIT in that order.
Each value is the median for one run of that build.

| Operation | Rows | JIT time | NativeAOT time | NativeAOT `Speed` time |
| --- | ---: | ---: | ---: | ---: |
| Load | 2,000 | 20.45 ms | 18.02 ms | 16.43 ms |
| Load | 20,000 | 132.22 ms | 205.91 ms | 191.35 ms |
| Save | 2,000 | 43.51 ms | 21.47 ms | 20.30 ms |
| Save | 20,000 | 184.78 ms | 236.59 ms | 222.98 ms |

`Speed` reduced the NativeAOT median in all four cases. JIT still had lower medians on
both 20,000-row operations. [NativeAOT optimization settings](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/optimizing)

NativeAOT avoids JIT compilation at process startup. In these warmed-up measurements,
the JIT can use runtime profiles to optimize frequently executed code paths. That may
help explain the results. A comparison on Linux would show whether the same pattern holds
there. [Tiered compilation and dynamic PGO](https://learn.microsoft.com/en-us/dotnet/core/runtime-config/compilation)

Run these commands from a Windows x64 Visual Studio Developer PowerShell with the C++
build tools installed.

```powershell
dotnet run --project tests/Aid.Xlsx.Benchmarks -c Release -- --export-input artifacts/native-benchmarks/input
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -o artifacts/native-benchmarks/jit
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -r win-x64 -p:PublishAot=true -p:NuGetLockFilePath=obj/packages.aot.lock.json -o artifacts/native-benchmarks/aot
dotnet publish tests/Aid.Xlsx.NativeBenchmarks -c Release -r win-x64 -p:PublishAot=true -p:OptimizationPreference=Speed -p:NuGetLockFilePath=obj/packages.aot.lock.json -o artifacts/native-benchmarks/aot-speed
.\artifacts\native-benchmarks\aot\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
.\artifacts\native-benchmarks\aot-speed\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
.\artifacts\native-benchmarks\jit\Aid.Xlsx.NativeBenchmarks.exe artifacts/native-benchmarks/input
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
