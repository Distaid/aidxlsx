using Aid.Xlsx.Workbooks;

var current = ApiSurface.Read(typeof(XlsxWorkbook).Assembly);
if (args is ["--print-api"])
{
    Console.Write(string.Join('\n', current.Order(StringComparer.Ordinal)) + "\n");
    return 0;
}

var baseline = File.ReadLines(Path.Combine(AppContext.BaseDirectory, "Baselines", "PublicApi.txt"))
    .ToHashSet(StringComparer.Ordinal);
var added = current.Except(baseline).Order(StringComparer.Ordinal).ToArray();
var removed = baseline.Except(current).Order(StringComparer.Ordinal).ToArray();
foreach (var line in added)
    Console.Error.WriteLine($"Added API declaration {line}");
foreach (var line in removed)
    Console.Error.WriteLine($"Removed API declaration {line}");

var invalid = current.Where(line => line.Contains("DocumentFormat.OpenXml.", StringComparison.Ordinal))
    .ToArray();
foreach (var line in invalid)
    Console.Error.WriteLine($"Unexpected public API {line}");

Console.WriteLine($"Compared {current.Count} Aid.Xlsx API declarations with {baseline.Count} checked-in declarations.");
var outputDirectory = args is ["--save-corpus", var directory] ? directory : null;
var corpusPassed = WorkbookCorpus.Verify(outputDirectory);
return corpusPassed && added.Length + removed.Length + invalid.Length == 0 ? 0 : 1;
