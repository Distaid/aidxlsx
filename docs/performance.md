# Workbook I/O performance

The standalone benchmark compares Aid.Xlsx with ClosedXML 0.105.1. Run its project directly. Both libraries load the same XLSX and save an equivalent workbook to a new memory stream. Source workbooks are loaded before each save measurement. The input contains one worksheet with five columns of numbers, repeated text, dates, and formulas. Validation and formula calculation are disabled during saving.

```sh
dotnet run --project tests/Aid.Xlsx.Benchmarks -c Release -- --filter '*' --artifacts artifacts/benchmarks
```

## Windows sample

BenchmarkDotNet 0.15.8 ShortRun, in-process, on Windows 11 with an AMD Ryzen 5 5600 and .NET 10.0.14. Times are means and allocations are managed bytes per operation. The 2,000-row timings varied between iterations.

| Operation | Rows | ClosedXML time | Aid.Xlsx time | ClosedXML allocated | Aid.Xlsx allocated |
| --- | ---: | ---: | ---: | ---: | ---: |
| Load | 2,000 | 32.40 ms | 11.45 ms | 8.53 MB | 3.74 MB |
| Load | 20,000 | 218.35 ms | 131.88 ms | 79.91 MB | 34.19 MB |
| Save | 2,000 | 35.41 ms | 33.53 ms | 11.31 MB | 5.65 MB |
| Save | 20,000 | 425.23 ms | 162.45 ms | 105.92 MB | 47.77 MB |
