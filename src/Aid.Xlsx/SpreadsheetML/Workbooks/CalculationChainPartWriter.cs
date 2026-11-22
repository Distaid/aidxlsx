using System.Xml.Linq;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.SpreadsheetML;

internal static class CalculationChainPartWriter
{
    private static readonly XNamespace Spreadsheet = OpenXmlConst.Main2006SsNs;

    internal static XDocument? GenerateContent(Stream? existing, XlsxWorkbook workbook)
    {
        var document = existing is null ? null : XDocumentExtensions.Load(existing);
        if (document?.Root is not { } root || root.Name != Spreadsheet + "calcChain")
        {
            root = new XElement(Spreadsheet + "calcChain");
            document = new XDocument(root);
        }

        if ((string?)root.Attribute("xmlns") == Spreadsheet.NamespaceName)
            root.Attribute("xmlns")!.Remove();
        root.SetAttributeValue(XNamespace.Xmlns + "x", Spreadsheet.NamespaceName);
        root.Elements(Spreadsheet + "c").Remove();
        var cells = workbook.WorksheetsInternal.Cast<XlsxWorksheet>().SelectMany(GetCalculationCells);
        root.AddFirst(cells);
        return root.HasElements ? document : null;
    }

    private static IEnumerable<XElement> GetCalculationCells(XlsxWorksheet sheet)
    {
        foreach (var cell in sheet.Internals.CellsCollection.GetCells(c => c.HasFormula))
        {
            // Excel does not generate calculation-chain entries for data tables.
            if (cell.Formula.Type == FormulaType.DataTable)
                continue;

            if (!cell.HasArrayFormula)
            {
                yield return CreateEntry(cell, sheet.SheetId);
                continue;
            }

            cell.FormulaReference ??= cell.AsRange().RangeAddress;
            if (!cell.FormulaReference.FirstAddress.Equals(cell.Address))
                continue;

            var anchor = CreateEntry(cell, sheet.SheetId);
            anchor.Add(new XAttribute("a", "1"));
            yield return anchor;

            foreach (var child in sheet.Range(cell.FormulaReference).Cells())
                yield return CreateEntry(child, sheet.SheetId);
        }
    }

    private static XElement CreateEntry(IXlsxCell cell, uint sheetId) =>
        new(Spreadsheet + "c", new XAttribute("r", cell.Address.ToString()!), new XAttribute("i", sheetId));
}
