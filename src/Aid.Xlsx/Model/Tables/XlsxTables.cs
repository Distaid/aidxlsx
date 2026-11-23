#nullable disable

using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.Tables;

using System.Collections;

internal class XlsxTables : IXlsxTables, IEnumerable<XlsxTable>
{
    private readonly Dictionary<string, XlsxTable> _tables;

    public XlsxTables()
    {
        _tables = [with(StringComparer.OrdinalIgnoreCase)];
        Deleted = new HashSet<string>();
    }

    internal ICollection<string> Deleted { get; }

    #region IXlsxTables Members

    bool IXlsxTables.TryGetTable(string tableName, out IXlsxTable table)
    {
        if (TryGetTable(tableName, out XlsxTable foundTable))
        {
            table = foundTable;
            return true;
        }

        table = default;
        return false;
    }

    public void Add(IXlsxTable table)
    {
        var xlTable = (XlsxTable)table;
        _tables.Add(table.Name, xlTable);
        xlTable.OnAddedToTables();
    }

    public IXlsxTables Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        _tables.Values.ForEach(t => t.Clear(clearOptions));
        return this;
    }

    public bool Contains(string name)
    {
        return _tables.ContainsKey(name);
    }

    public Dictionary<string, XlsxTable>.ValueCollection.Enumerator GetEnumerator()
    {
        return _tables.Values.GetEnumerator();
    }

    IEnumerator<XlsxTable> IEnumerable<XlsxTable>.GetEnumerator() => GetEnumerator();

    IEnumerator<IXlsxTable> IEnumerable<IXlsxTable>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public void Remove(int index)
    {
        this.Remove(_tables.ElementAt(index).Key);
    }

    public void Remove(string name)
    {
        if (!_tables.TryGetValue(name, out XlsxTable table))
            throw new ArgumentOutOfRangeException(nameof(name), $"Unable to delete table because the table name {name} could not be found.");

        _tables.Remove(name);

        var relId = table.RelId;

        if (relId is not null)
            Deleted.Add(relId);
    }

    public IXlsxTable Table(int index)
    {
        return _tables.ElementAt(index).Value;
    }

    public IXlsxTable Table(string name)
    {
        if (TryGetTable(name, out XlsxTable table))
            return table;

        throw new ArgumentOutOfRangeException(nameof(name), $"Table {name} was not found.");
    }

    internal bool TryGetTable(string tableName, out XlsxTable table)
    {
        return _tables.TryGetValue(tableName, out table);
    }

    #endregion IXlsxTables Members
}
