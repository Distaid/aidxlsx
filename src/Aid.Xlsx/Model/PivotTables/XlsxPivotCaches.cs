using System.Collections;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.PivotTables;

internal class XlsxPivotCaches : IXlsxPivotCaches, IEnumerable<XlsxPivotCache>
{
    private readonly XlsxWorkbook _workbook;
    private readonly List<XlsxPivotCache> _caches = new();

    public XlsxPivotCaches(XlsxWorkbook workbook)
    {
        _workbook = workbook;
    }

    IXlsxPivotCache IXlsxPivotCaches.Add(IXlsxRange range) => Add(SheetArea.From(range));

    IEnumerator<IXlsxPivotCache> IEnumerable<IXlsxPivotCache>.GetEnumerator() => GetEnumerator();

    IEnumerator<XlsxPivotCache> IEnumerable<XlsxPivotCache>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public List<XlsxPivotCache>.Enumerator GetEnumerator() => _caches.GetEnumerator();

    internal XlsxPivotCache Add(SheetArea area)
    {
        var source = _workbook.TryGetTable(area, out var table)
            ? new XlsxPivotSourceReference(table.Name)
            : new XlsxPivotSourceReference(area);

        var newPivotCache = new XlsxPivotCache(source, _workbook);
        newPivotCache.Refresh();
        _caches.Add(newPivotCache);
        return newPivotCache;
    }

    internal XlsxPivotCache Add(IXlsxPivotSource source)
    {
        var newPivotCache = new XlsxPivotCache(source, _workbook);
        _caches.Add(newPivotCache);
        return newPivotCache;
    }

    /// <summary>
    /// Try to find an existing pivot cache for the passed area. The area
    /// is checked against both types of source references (tables and
    /// ranges) and if area matches, the cache is returned.
    /// </summary>
    internal XlsxPivotCache? Find(SheetArea area)
    {
        // This method mimics behavior of Excel.
        // If there is a table for the area and there is a cache for the table, return cache for the table.
        if (_workbook.TryGetTable(area, out var table))
        {
            // Table exists, so try to find it and match with the source reference.
            var tableSource = new XlsxPivotSourceReference(table.Name);
            foreach (var cache in _caches)
            {
                if (cache.Source.Equals(tableSource))
                    return cache;
            }
        }

        // Try to find a cache with area source.
        var areaSource = new XlsxPivotSourceReference(area);
        foreach (var cache in _caches)
        {
            if (cache.Source.Equals(areaSource))
                return cache;
        }

        return null;
    }
}
