#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.Tables;

public interface IXlsxTables : IEnumerable<IXlsxTable>
{
    void Add(IXlsxTable table);

    /// <summary>
    /// Clears the contents of these tables.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    IXlsxTables Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    bool Contains(string name);

    void Remove(int index);

    void Remove(string name);

    IXlsxTable Table(int index);

    IXlsxTable Table(string name);

    bool TryGetTable(string tableName, out IXlsxTable table);
}
