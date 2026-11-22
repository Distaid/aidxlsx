using Aid.Xlsx.Workbooks;

var current = ApiSurface.Read(typeof(XlsxWorkbook).Assembly);
var upstream = ApiSurface.Read(typeof(ClosedXML.Excel.XLWorkbook).Assembly);
var differences = upstream.Except(current).Select(line => $"- {line}")
    .Concat(current.Except(upstream).Select(line => $"+ {line}"))
    .Order(StringComparer.Ordinal)
    .ToArray();

if (args is ["--print-differences"])
{
    foreach (var line in differences)
        Console.WriteLine(line);
    return 0;
}

var accepted = File.ReadLines(Path.Combine(AppContext.BaseDirectory, "AcceptedDifferences.txt"))
    .Where(line => line.StartsWith("+ ", StringComparison.Ordinal) || line.StartsWith("- ", StringComparison.Ordinal))
    .ToHashSet(StringComparer.Ordinal);
var unexpected = differences.Where(line => !accepted.Contains(line)).ToArray();
var resolved = accepted.Except(differences).Order(StringComparer.Ordinal).ToArray();
foreach (var line in unexpected)
    Console.Error.WriteLine($"Unexpected API difference {line}");
foreach (var line in resolved)
    Console.Error.WriteLine($"Resolved API difference {line}");

var invalid = typeof(XlsxWorkbook).Assembly.GetExportedTypes()
    .Where(type => type.Namespace == "Aid.Xlsx" || ApiSurface.Normalize(type.Name) != type.Name)
    .Select(type => type.FullName)
    .Concat(current.Where(line => line.Contains("DocumentFormat.OpenXml.", StringComparison.Ordinal)))
    .ToArray();
foreach (var line in invalid)
    Console.Error.WriteLine($"Unexpected public API {line}");

Console.WriteLine($"Compared {current.Count} Aid.Xlsx API declarations with {upstream.Count} upstream declarations. Accepted differences {accepted.Count}.");
var outputDirectory = args is ["--save-corpus", var directory] ? directory : null;
var corpusPassed = WorkbookCorpus.Verify(outputDirectory);
return corpusPassed && unexpected.Length + resolved.Length + invalid.Length == 0 ? 0 : 1;
