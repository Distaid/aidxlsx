#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Ranges.Sort;
using Aid.Xlsx.Ranges;

namespace Aid.Xlsx.Columns
{
    public interface IXlsxColumn : IXlsxRangeBase
    {
        /// <summary>
        /// Gets or sets the width of this column in number of characters (NoC).
        /// </summary>
        /// <remarks>
        /// NoC are a non-linear units displayed as a column width in Excel, next to pixels. NoC combined with default font
        /// of the workbook can express width of the column in pixels and other units.
        /// </remarks>
        Double Width { get; set; }

        /// <summary>
        /// Deletes this column and shifts the columns at the right of this one accordingly.
        /// </summary>
        /// <remarks>Don't use in a loop due to poor performance. Use <see cref="IXlsxRange.Delete(XlsxShiftDeletedCells)"/> instead.</remarks>
        void Delete();

        /// <summary>
        /// Gets this column's number
        /// </summary>
        Int32 ColumnNumber();

        /// <summary>
        /// Gets this column's letter
        /// </summary>
        String ColumnLetter();

        /// <summary>
        /// Inserts X number of columns at the right of this one.
        /// <para>All columns at the right will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfColumns">The number of columns to insert.</param>
        IXlsxColumns InsertColumnsAfter(Int32 numberOfColumns);

        /// <summary>
        /// Inserts X number of columns at the left of this one.
        /// <para>This column and all at the right will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfColumns">The number of columns to insert.</param>
        IXlsxColumns InsertColumnsBefore(Int32 numberOfColumns);

        /// <summary>
        /// Gets the cell in the specified row.
        /// </summary>
        /// <param name="rowNumber">The cell's row.</param>
        IXlsxCell Cell(Int32 rowNumber);

        /// <summary>
        /// Returns the specified group of cells, separated by commas.
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
        /// </summary>
        /// <param name="cellsInColumn">The column cells to return.</param>
        new IXlsxCells Cells(String cellsInColumn);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstRow">The first row in the group of cells to return.</param>
        /// <param name="lastRow">The last row in the group of cells to return.</param>
        IXlsxCells Cells(Int32 firstRow, Int32 lastRow);

        /// <summary>
        /// Adjusts the width of the column based on its contents.
        /// </summary>
        IXlsxColumn AdjustToContents();

        /// <summary>
        /// Adjusts the width of the column based on its contents, starting from the startRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        IXlsxColumn AdjustToContents(Int32 startRow);

        /// <summary>
        /// Adjusts the width of the column based on its contents, starting from the startRow and ending at endRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        /// <param name="endRow">The row to end calculating the column width.</param>
        IXlsxColumn AdjustToContents(Int32 startRow, Int32 endRow);

        IXlsxColumn AdjustToContents(Double minWidth, Double maxWidth);

        IXlsxColumn AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth);

        /// <summary>
        /// Adjust width of the column according to the content of the cells.
        /// </summary>
        /// <param name="startRow">Number of a first row whose content is considered.</param>
        /// <param name="endRow">Number of a last row whose content is considered.</param>
        /// <param name="minWidth">Minimum width of adjusted column, in NoC.</param>
        /// <param name="maxWidth">Maximum width of adjusted column, in NoC.</param>
        IXlsxColumn AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth);

        /// <summary>
        /// Hides this column.
        /// </summary>
        IXlsxColumn Hide();

        /// <summary>Unhides this column.</summary>
        IXlsxColumn Unhide();

        /// <summary>
        /// Gets a value indicating whether this column is hidden or not.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this column is hidden; otherwise, <c>false</c>.
        /// </value>
        Boolean IsHidden { get; }

        /// <summary>
        /// Gets or sets the outline level of this column.
        /// </summary>
        /// <value>
        /// The outline level of this column.
        /// </value>
        Int32 OutlineLevel { get; set; }

        /// <summary>
        /// Adds this column to the next outline level (Increments the outline level for this column by 1).
        /// </summary>
        IXlsxColumn Group();

        /// <summary>
        /// Adds this column to the next outline level (Increments the outline level for this column by 1).
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
        IXlsxColumn Group(Boolean collapse);

        /// <summary>
        /// Sets outline level for this column.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXlsxColumn Group(Int32 outlineLevel);

        /// <summary>
        /// Sets outline level for this column.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
        IXlsxColumn Group(Int32 outlineLevel, Boolean collapse);

        /// <summary>
        /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
        /// </summary>
        IXlsxColumn Ungroup();

        /// <summary>
        /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove this column from all outline levels.</param>
        IXlsxColumn Ungroup(Boolean fromAll);

        /// <summary>
        /// Show this column as collapsed.
        /// </summary>
        IXlsxColumn Collapse();

        /// <summary>Expands this column (if it's collapsed).</summary>
        IXlsxColumn Expand();

        Int32 CellCount();

        IXlsxRangeColumn CopyTo(IXlsxCell cell);

        IXlsxRangeColumn CopyTo(IXlsxRangeBase range);

        IXlsxColumn CopyTo(IXlsxColumn column);

        IXlsxColumn Sort(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXlsxRangeColumn Column(Int32 start, Int32 end);

        IXlsxRangeColumn Column(IXlsxCell start, IXlsxCell end);

        IXlsxRangeColumns Columns(String columns);

        /// <summary>
        /// Adds a vertical page break after this column.
        /// </summary>
        IXlsxColumn AddVerticalPageBreak();

        IXlsxColumn ColumnLeft();

        IXlsxColumn ColumnLeft(Int32 step);

        IXlsxColumn ColumnRight();

        IXlsxColumn ColumnRight(Int32 step);

        /// <summary>
        /// Clears the contents of this column.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXlsxColumn Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

        IXlsxRangeColumn ColumnUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents);
    }
}
