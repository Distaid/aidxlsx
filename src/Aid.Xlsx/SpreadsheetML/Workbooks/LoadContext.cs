using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Opc.Xml;

namespace Aid.Xlsx.SpreadsheetML.Workbooks;

internal class LoadContext
{
    /// <summary>
    /// Conditional formats for pivot tables, loaded from sheets. Key is sheet name, value is the
    /// conditional formats.
    /// </summary>
    private readonly Dictionary<string, List<XlsxConditionalFormat>> _pivotCfs = new(XlsxHelper.SheetComparer);

    internal void AddPivotTableCf(string sheetName, XlsxConditionalFormat conditionalFormat)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out List<XlsxConditionalFormat>? list))
        {
            list = [];
            _pivotCfs[sheetName] = list;
        }

        list.Add(conditionalFormat);
    }

    internal XlsxConditionalFormat GetPivotCf(string sheetName, int priority)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out List<XlsxConditionalFormat>? list))
            throw PivotCfNotFoundException(sheetName, priority);

        XlsxConditionalFormat? pivotCf = list.SingleOrDefault(x => x.Priority == priority) ?? throw PivotCfNotFoundException(sheetName, priority);
        return pivotCf;
    }

    private static Exception PivotCfNotFoundException(string sheetName, int priority)
    {
        return PartStructureException.ExpectedElementNotFound($"conditional formatting for pivot table in sheet {sheetName} with priority {priority}");
    }
}
