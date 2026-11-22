#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Tables;

namespace Aid.Xlsx.PivotTables
{
    public interface IXlsxPivotTables : IEnumerable<IXlsxPivotTable>
    {
        /// <summary>
        /// Add a pivot table that will use the pivot cache.
        /// </summary>
        /// <param name="name">Name of new pivot table.</param>
        /// <param name="targetCell">A cell where will the pivot table be have it's left top corner.</param>
        /// <param name="pivotCache">Pivot cache to use for the pivot table.</param>
        /// <returns>Added pivot table.</returns>
        /// <exception cref="ArgumentException">There already is a pivot table with the same name.</exception>
        IXlsxPivotTable Add(String name, IXlsxCell targetCell, IXlsxPivotCache pivotCache);

        /// <summary>
        /// Add a pivot table from source data of <paramref name="range"/>.
        /// If workbook already contains a cache for same range as the
        /// <paramref name="range"/>, the matching pivot cache is used.
        /// </summary>
        /// <param name="name">Name of new pivot table</param>
        /// <param name="targetCell">A cell where will the pivot table be have it's left top corner.</param>
        /// <param name="range">A range to add/find pivot cache.</param>
        /// <exception cref="ArgumentException">There already is a pivot table with the same name.</exception>
        IXlsxPivotTable Add(String name, IXlsxCell targetCell, IXlsxRange range);

        /// <summary>
        /// Add a pivot table from source data of <paramref name="table"/>.
        /// If workbook already contains a cache for same range as the
        /// <paramref name="table"/>, the matching pivot cache is used.
        /// </summary>
        /// <param name="name">Name of new pivot table</param>
        /// <param name="targetCell">A cell where will the pivot table be have it's left top corner.</param>
        /// <param name="table">A table to add/find pivot cache.</param>
        /// <exception cref="ArgumentException">There already is a pivot table with the same name.</exception>
        IXlsxPivotTable Add(String name, IXlsxCell targetCell, IXlsxTable table);

        Boolean Contains(String name);

        void Delete(String name);

        void DeleteAll();

        /// <summary>
        /// Get pivot table with the specified name (case insensitive).
        /// </summary>
        /// <param name="name">Name of a pivot table to return.</param>
        /// <exception cref="KeyNotFoundException">No such pivot table found.</exception>
        IXlsxPivotTable PivotTable(String name);
    }
}
