#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Tables;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.PivotTables
{
    internal class XlsxPivotTables : IXlsxPivotTables, IEnumerable<XlsxPivotTable>
    {
        private readonly Dictionary<String, XlsxPivotTable> _pivotTables = new(StringComparer.OrdinalIgnoreCase);

        public XlsxPivotTables(XlsxWorksheet worksheet)
        {
            Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        internal XlsxWorksheet Worksheet { get; }

        public void Add(XlsxPivotTable pivotTable)
        {
            var pivotCache = pivotTable.PivotCache;
            if (!pivotCache.FieldNames.Any())
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
            pivotTable.UpdateCacheFields(Array.Empty<string>());
            return pivotTable;
        }

        public IXlsxPivotTable Add(string name, IXlsxCell targetCell, IXlsxRange range)
        {
            var area = SheetArea.From(range);
            var pivotCaches = Worksheet.Workbook.PivotCachesInternal;
            var existingPivotCache = pivotCaches.Find(area);
            var pivotCache = existingPivotCache ?? pivotCaches.Add(area);
            return Add(name, targetCell, pivotCache);
        }

        public IXlsxPivotTable Add(string name, IXlsxCell targetCell, IXlsxTable table)
        {
            return Add(name, targetCell, (IXlsxRange)table);
        }

        public Boolean Contains(String name)
        {
            return _pivotTables.ContainsKey(name);
        }

        public void Delete(String name)
        {
            _pivotTables.Remove(name);
        }

        public void DeleteAll()
        {
            _pivotTables.Clear();
        }

        IXlsxPivotTable IXlsxPivotTables.PivotTable(String name)
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

        internal void Add(String name, IXlsxPivotTable pivotTable)
        {
            _pivotTables.Add(name, (XlsxPivotTable)pivotTable);
        }

        /// <inheritdoc cref="IXlsxPivotTables.PivotTable"/>
        internal XlsxPivotTable PivotTable(String name)
        {
            return _pivotTables[name];
        }
    }
}
