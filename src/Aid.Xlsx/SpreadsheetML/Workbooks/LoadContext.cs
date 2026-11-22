using Aid.Xlsx.ConditionalFormats;
using Aid.Xlsx.Opc;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.SpreadsheetML;

internal class LoadContext
{
    /// <summary>
    /// Conditional formats for pivot tables, loaded from sheets. Key is sheet name, value is the
    /// conditional formats.
    /// </summary>
    private readonly Dictionary<string, List<XlsxConditionalFormat>> _pivotCfs = new(XlsxHelper.SheetComparer);

    internal void AddPivotTableCf(string sheetName, XlsxConditionalFormat conditionalFormat)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out var list))
        {
            list = new List<XlsxConditionalFormat>();
            _pivotCfs[sheetName] = list;
        }

        list.Add(conditionalFormat);
    }

    internal XlsxConditionalFormat GetPivotCf(string sheetName, int priority)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out var list))
            throw PivotCfNotFoundException(sheetName, priority);

        var pivotCf = list.SingleOrDefault(x => x.Priority == priority);
        if (pivotCf is null)
            throw PivotCfNotFoundException(sheetName, priority);

        return pivotCf;
    }

    private static Exception PivotCfNotFoundException(string sheetName, int priority)
    {
        return PartStructureException.ExpectedElementNotFound($"conditional formatting for pivot table in sheet {sheetName} with priority {priority}");
    }
}
