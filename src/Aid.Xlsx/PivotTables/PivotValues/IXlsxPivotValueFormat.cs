using Aid.Xlsx.Styles;

namespace Aid.Xlsx.PivotTables.PivotValues
{
    /// <summary>
    /// An API for manipulating a <see cref="IXlsxPivotValue.NumberFormat">format</see> of one
    /// <see cref="IXlsxPivotTable"/> <see cref="IXlsxPivotValue">data field</see>.
    /// </summary>
    public interface IXlsxPivotValueFormat : IXlsxNumberFormatBase
    {
        /// <summary>
        /// Set number formatting using one of predefined codes. Predefined codes are described in
        /// the <see cref="XlsxPredefinedFormat"/>.
        /// </summary>
        /// <param name="value">A numeric value describing how should the number be formatted.</param>
        /// <exception cref="ArgumentOutOfRangeException">The passed value is not a predefined format.</exception>
        IXlsxPivotValue SetNumberFormatId(Int32 value);

        IXlsxPivotValue SetFormat(String value);
    }
}
