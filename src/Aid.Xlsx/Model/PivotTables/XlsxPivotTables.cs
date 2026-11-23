#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PivotTables;

internal class XlsxPivotTables(XlsxWorksheet worksheet) : IXlsxPivotTables, IEnumerable<XlsxPivotTable>
{
    private readonly Dictionary<string, XlsxPivotTable> _pivotTables = [with(StringComparer.OrdinalIgnoreCase)];

    internal XlsxWorksheet Worksheet { get; } = worksheet ?? throw new ArgumentNullException(nameof(worksheet));

    public void Add(XlsxPivotTable pivotTable)
    {
        XlsxPivotCache pivotCache = pivotTable.PivotCache;
        if (pivotCache.FieldNames.Count == 0)
            pivotCache.Refresh();

        _pivotTables.Add(pivotTable.Name, pivotTable);
    }

    public IXlsxPivotTable Add(string name, IXlsxCell targetCell, IXlsxPivotCache pivotCache)
    {
        var pivotTable = new XlsxPivotTable(Worksheet, (XlsxPivotCache)pivotCache)
        {
            Name = name,
            Area = new Area(Point.FromAddress(targetCell.Address)),
        };
        Add(pivotTable);
        pivotTable.UpdateCacheFields([]);
        return pivotTable;
    }

    public IXlsxPivotTable Add(string name, IXlsxCell targetCell, IXlsxRange range)
    {
        var area = SheetArea.From(range);
        XlsxPivotCaches pivotCaches = Worksheet.Workbook.PivotCachesInternal;
        XlsxPivotCache existingPivotCache = pivotCaches.Find(area);
        XlsxPivotCache pivotCache = existingPivotCache ?? pivotCaches.Add(area);
        return Add(name, targetCell, pivotCache);
    }

    public IXlsxPivotTable Add(string name, IXlsxCell targetCell, IXlsxTable table)
    {
        return Add(name, targetCell, (IXlsxRange)table);
    }

    public bool Contains(string name)
    {
        return _pivotTables.ContainsKey(name);
    }

    public void Delete(string name)
    {
        _pivotTables.Remove(name);
    }

    public void DeleteAll()
    {
        _pivotTables.Clear();
    }

    IXlsxPivotTable IXlsxPivotTables.PivotTable(string name)
    {
        return PivotTable(name);
    }

    IEnumerator<IXlsxPivotTable> IEnumerable<IXlsxPivotTable>.GetEnumerator()
    {
        return GetEnumerator();
    }

    IEnumerator<XlsxPivotTable> IEnumerable<XlsxPivotTable>.GetEnumerator()
    {
        return GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public Dictionary<string, XlsxPivotTable>.ValueCollection.Enumerator GetEnumerator()
    {
        return _pivotTables.Values.GetEnumerator();
    }

    internal void Add(string name, IXlsxPivotTable pivotTable)
    {
        _pivotTables.Add(name, (XlsxPivotTable)pivotTable);
    }

    /// <inheritdoc cref="IXlsxPivotTables.PivotTable"/>
    internal XlsxPivotTable PivotTable(string name)
    {
        return _pivotTables[name];
    }
}
