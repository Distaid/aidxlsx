#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Cells;

namespace Aid.Xlsx.Tables
{
    public interface IXlsxTables : IEnumerable<IXlsxTable>
    {
        void Add(IXlsxTable table);

        /// <summary>
        /// Clears the contents of these tables.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXlsxTables Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

        Boolean Contains(String name);

        void Remove(Int32 index);

        void Remove(String name);

        IXlsxTable Table(Int32 index);

        IXlsxTable Table(String name);

        Boolean TryGetTable(String tableName, out IXlsxTable table);
    }
}
