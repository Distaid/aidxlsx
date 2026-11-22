#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Diagnostics;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.Drawings;

[DebuggerDisplay("{Address} {Offset}")]
internal class XlsxMarker
{
    // Using a range to store the location so that it gets added to the range repository
    // and hence will be adjusted when there are insertions / deletions
    private readonly IXlsxRange rangeCell;

    internal XlsxMarker(IXlsxCell cell)
                : this(cell.AsRange(), new System.Drawing.Point(0, 0))
    { }

    internal XlsxMarker(IXlsxCell cell, System.Drawing.Point offset)
        : this(cell.AsRange(), offset)
    { }

    private XlsxMarker(IXlsxRange rangeCell, System.Drawing.Point offset)
    {
        if (rangeCell.RowCount() != 1 || rangeCell.ColumnCount() != 1)
            throw new ArgumentException("Range should contain only one cell.", nameof(rangeCell));

        this.rangeCell = rangeCell;
        this.Offset = offset;
    }

    public IXlsxCell Cell { get => rangeCell.FirstCell(); }
    public Int32 ColumnNumber { get => rangeCell.RangeAddress.FirstAddress.ColumnNumber; }
    public System.Drawing.Point Offset { get; set; }
    public Int32 RowNumber { get => rangeCell.RangeAddress.FirstAddress.RowNumber; }
}
