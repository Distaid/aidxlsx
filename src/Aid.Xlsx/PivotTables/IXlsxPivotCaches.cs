using Aid.Xlsx.Ranges;
using Aid.Xlsx.Tables;

namespace Aid.Xlsx.PivotTables
{
    /// <summary>
    /// A collection of <see cref="IXlsxPivotCache">pivot caches</see>. Pivot cache
    /// can be added from a <see cref="IXlsxRange"/> or a <see cref="IXlsxTable"/>.
    /// </summary>
    public interface IXlsxPivotCaches : IEnumerable<IXlsxPivotCache>
    {
        /// <summary>
        /// Add a new pivot cache for the range. If the range area is same as
        /// an area of a table, the created cache will reference the table
        /// as source of data instead of a range of cells.
        /// </summary>
        /// <param name="range">Range for which to create the pivot cache.</param>
        /// <returns>The pivot cache for the range.</returns>
        IXlsxPivotCache Add(IXlsxRange range);
    }
}
