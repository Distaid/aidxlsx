#nullable disable

using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Coordinates
{
    /// <summary>
    /// Reference to a single cell in a workbook. Reference can be absolute, relative or mixed.
    /// Reference can be with or without a worksheet.
    /// </summary>
    public interface IXlsxAddress : IEqualityComparer<IXlsxAddress>, IEquatable<IXlsxAddress>
    {
        String ColumnLetter { get; }
        Int32 ColumnNumber { get; }
        Boolean FixedColumn { get; }
        Boolean FixedRow { get; }
        Int32 RowNumber { get; }
        String UniqueId { get; }

        /// <summary>
        /// Worksheet of the reference. Value is null for address without a worksheet.
        /// </summary>
        IXlsxWorksheet Worksheet { get; }

        String ToString(XlsxReferenceStyle referenceStyle);

        String ToString(XlsxReferenceStyle referenceStyle, Boolean includeSheet);

        String ToStringFixed();

        String ToStringFixed(XlsxReferenceStyle referenceStyle);

        String ToStringFixed(XlsxReferenceStyle referenceStyle, Boolean includeSheet);

        String ToStringRelative();

        String ToStringRelative(Boolean includeSheet);
    }
}
