#nullable disable

using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Coordinates;

/// <summary>
/// Reference to a single cell in a workbook. Reference can be absolute, relative or mixed.
/// Reference can be with or without a worksheet.
/// </summary>
public interface IXlsxAddress : IEqualityComparer<IXlsxAddress>, IEquatable<IXlsxAddress>
{
    string ColumnLetter { get; }
    int ColumnNumber { get; }
    bool FixedColumn { get; }
    bool FixedRow { get; }
    int RowNumber { get; }
    string UniqueId { get; }

    /// <summary>
    /// Worksheet of the reference. Value is null for address without a worksheet.
    /// </summary>
    IXlsxWorksheet Worksheet { get; }

    string ToString(XlsxReferenceStyle referenceStyle);

    string ToString(XlsxReferenceStyle referenceStyle, bool includeSheet);

    string ToStringFixed();

    string ToStringFixed(XlsxReferenceStyle referenceStyle);

    string ToStringFixed(XlsxReferenceStyle referenceStyle, bool includeSheet);

    string ToStringRelative();

    string ToStringRelative(bool includeSheet);
}
