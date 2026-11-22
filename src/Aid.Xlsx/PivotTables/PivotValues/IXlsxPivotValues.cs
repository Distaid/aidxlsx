#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.PivotTables.Areas;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.PivotTables.PivotValues
{
    public interface IXlsxPivotValues : IEnumerable<IXlsxPivotValue>
    {
        /// <summary>
        /// Add a new value field to the pivot table. If addition would cause, the
        /// <see cref="XlsxConstants.PivotTable.ValuesSentinalLabel"/> field is added to the
        /// <see cref="IXlsxPivotTable.ColumnLabels"/>. The added field will use passed
        /// <paramref name="sourceName"/> as the <see cref="IXlsxPivotField.CustomName"/>.
        /// </summary>
        /// <param name="sourceName">The <see cref="IXlsxPivotField.SourceName"/> that is used as a
        ///     data. Multiple data fields can use same source (e.g. sum and count).</param>
        /// <returns>Newly added field.</returns>
        IXlsxPivotValue Add(String sourceName);

        /// <summary>
        /// Add a new value field to the pivot table. If addition would cause, the
        /// <see cref="XlsxConstants.PivotTable.ValuesSentinalLabel"/> field is added to the
        /// <see cref="IXlsxPivotTable.ColumnLabels"/>.
        /// </summary>
        /// <param name="sourceName">The <see cref="IXlsxPivotField.SourceName"/> that is used as a
        ///     data. Multiple data fields can use same source (e.g. sum and count).</param>
        /// <param name="customName">The added data field <see cref="IXlsxPivotField.CustomName"/>.</param>
        /// <returns>Newly added field.</returns>
        IXlsxPivotValue Add(String sourceName, String customName);

        void Clear();

        Boolean Contains(String customName);

        Boolean Contains(IXlsxPivotValue pivotValue);

        IXlsxPivotValue Get(String customName);

        IXlsxPivotValue Get(Int32 index);

        Int32 IndexOf(String customName);

        Int32 IndexOf(IXlsxPivotValue pivotValue);

        void Remove(String customName);
    }
}
