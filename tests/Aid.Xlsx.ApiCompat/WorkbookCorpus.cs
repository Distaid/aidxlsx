using System.Globalization;
using System.Text.Json;
using Aid.Xlsx.Workbooks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Upstream = ClosedXML.Excel;

internal static class WorkbookCorpus
{
    internal static bool Verify(string? outputDirectory)
    {
        var fixtures = Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Corpus"));
        if (fixtures.Length != 10)
            throw new InvalidOperationException($"Expected ten workbook fixtures, found {fixtures.Length}.");
        if (outputDirectory is not null)
            Directory.CreateDirectory(outputDirectory);

        var passed = true;
        foreach (var fixture in fixtures.Order(StringComparer.Ordinal))
        {
            var fileName = Path.GetFileName(fixture);
            using var baseline = new Upstream.XLWorkbook(fixture);
            using var workbook = new XlsxWorkbook(fixture);
            var expected = Capture(baseline);
            passed &= Compare(fileName, "load", expected, Capture(workbook));

            // The default save omits formula caches and leaves evaluation to the spreadsheet application.
            using var saved = new MemoryStream();
            workbook.SaveAs(saved, new SaveOptions { ValidatePackage = true });
            if (outputDirectory is not null)
            {
                // Use a distinct name so that spreadsheet applications can open both versions together.
                var outputName = $"{Path.GetFileNameWithoutExtension(fileName)}-roundtrip{Path.GetExtension(fileName)}";
                using var file = File.Create(Path.Combine(outputDirectory, outputName));
                saved.WriteTo(file);
            }
            saved.Position = 0;
            using var reopened = new Upstream.XLWorkbook(saved);
            passed &= Compare(fileName, "save", Capture(baseline, includeFormulaCache: false), Capture(reopened, includeFormulaCache: false));
            saved.Position = 0;
            using var document = SpreadsheetDocument.Open(saved, false);
            var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToArray();
            foreach (var error in errors)
                Console.Error.WriteLine($"{fileName} schema {error.Part?.Uri} {error.Path?.XPath}: {error.Description}");
            passed &= errors.Length == 0;
        }

        Console.WriteLine($"Compared load and save snapshots and validated outputs for {fixtures.Length} workbook fixtures.");
        return passed;
    }

    private static bool Compare(string fileName, string phase, string[] expected, string[] actual)
    {
        if (expected.SequenceEqual(actual, StringComparer.Ordinal))
            return true;

        foreach (var line in expected.Except(actual, StringComparer.Ordinal).Take(12))
            Console.Error.WriteLine($"{fileName} {phase} missing {line}");
        foreach (var line in actual.Except(expected, StringComparer.Ordinal).Take(12))
            Console.Error.WriteLine($"{fileName} {phase} unexpected {line}");
        return false;
    }

    private static string[] Capture(XlsxWorkbook workbook) => workbook.Worksheets.SelectMany(sheet =>
        new[] { Entry("sheet", sheet.Name, sheet.LastCellUsed()?.Address.ToString()) }
            .Concat(sheet.MergedRanges.Select(range => Entry("merge", range.RangeAddress.ToString())).Order(StringComparer.Ordinal))
            .Concat(sheet.Tables.Select(table => Entry("table", table.Name, table.RangeAddress.ToString(), table.Fields.Select(field => field.Name))).Order(StringComparer.Ordinal))
            .Concat(sheet.PivotTables.Select(pivot => Entry("pivot", pivot.Name)).Order(StringComparer.Ordinal))
            .Concat(sheet.CellsUsed().OrderBy(cell => cell.Address.RowNumber).ThenBy(cell => cell.Address.ColumnNumber).Select(cell =>
            {
                var link = cell.HasHyperlink ? cell.GetHyperlink() : null;
                var comment = cell.HasComment ? cell.GetComment() : null;
                return CellEntry(cell.Address.ToString(), cell.CachedValue.Type.ToString(), cell.CachedValue.ToString(CultureInfo.InvariantCulture),
                    cell.FormulaA1, cell.Style.NumberFormat.NumberFormatId, cell.Style.NumberFormat.Format,
                    cell.Style.Font.Bold, cell.Style.Font.Italic, cell.Style.Alignment.Horizontal.ToString(),
                    comment?.Author, comment?.ToString(), link?.IsExternal == true ? link.ExternalAddress.OriginalString : link?.InternalAddress);
            }))).ToArray();

    private static string[] Capture(Upstream.XLWorkbook workbook, bool includeFormulaCache = true) => workbook.Worksheets.SelectMany(sheet =>
        new[] { Entry("sheet", sheet.Name, sheet.LastCellUsed()?.Address.ToString()) }
            .Concat(sheet.MergedRanges.Select(range => Entry("merge", range.RangeAddress.ToString())).Order(StringComparer.Ordinal))
            .Concat(sheet.Tables.Select(table => Entry("table", table.Name, table.RangeAddress.ToString(), table.Fields.Select(field => field.Name))).Order(StringComparer.Ordinal))
            .Concat(sheet.PivotTables.Select(pivot => Entry("pivot", pivot.Name)).Order(StringComparer.Ordinal))
            .Concat(sheet.CellsUsed().OrderBy(cell => cell.Address.RowNumber).ThenBy(cell => cell.Address.ColumnNumber).Select(cell =>
            {
                var link = cell.HasHyperlink ? cell.GetHyperlink() : null;
                var comment = cell.HasComment ? cell.GetComment() : null;
                var value = includeFormulaCache || !cell.HasFormula ? cell.CachedValue : default;
                return CellEntry(cell.Address.ToString(), value.Type.ToString(), value.ToString(CultureInfo.InvariantCulture),
                    cell.FormulaA1, cell.Style.NumberFormat.NumberFormatId, cell.Style.NumberFormat.Format,
                    cell.Style.Font.Bold, cell.Style.Font.Italic, cell.Style.Alignment.Horizontal.ToString(),
                    comment?.Author, comment?.ToString(), link?.IsExternal == true ? link.ExternalAddress.OriginalString : link?.InternalAddress);
            }))).ToArray();

    private static string CellEntry(string? address, string type, string value, string formula, int formatId, string format,
        bool bold, bool italic, string alignment, string? author, string? comment, string? hyperlink)
    {
        // The format ID defines a built-in format and its display varies with locale.
        return Entry("cell", address, type, value, formula, formatId, formatId < 0 ? format : null, bold, italic,
            alignment, author, comment, hyperlink);
    }

    private static string Entry(params object?[] values) => JsonSerializer.Serialize(values);
}
