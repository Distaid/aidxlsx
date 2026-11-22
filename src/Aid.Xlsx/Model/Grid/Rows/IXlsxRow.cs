#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Ranges.Sort;
using Aid.Xlsx.Ranges;

namespace Aid.Xlsx.Rows
{
    public interface IXlsxRow : IXlsxRangeBase
    {
        /// <summary>
        /// Gets or sets the height of this row.
        /// </summary>
        /// <value>
        /// The width of this row in points.
        /// </value>
        Double Height { get; set; }

        /// <summary>
        /// Clears the height for the row and defaults it to the spreadsheet row height.
        /// </summary>
        void ClearHeight();

        /// <summary>
        /// Deletes this row and shifts the rows below this one accordingly.
        /// </summary>
        /// <remarks>Don't use in a loop due to poor performance. Use <see cref="IXlsxRange.Delete(XlsxShiftDeletedCells)"/> instead.</remarks>
        void Delete();

        /// <summary>
        /// Gets this row's number
        /// </summary>
        Int32 RowNumber();

        /// <summary>
        /// Inserts X number of rows below this one.
        /// <para>All rows below will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfRows">The number of rows to insert.</param>
        IXlsxRows InsertRowsBelow(Int32 numberOfRows);

        /// <summary>
        /// Inserts X number of rows above this one.
        /// <para>This row and all below will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfRows">The number of rows to insert.</param>
        IXlsxRows InsertRowsAbove(Int32 numberOfRows);

        IXlsxRow AdjustToContents();

        /// <summary>
        /// Adjusts the height of the row based on its contents, starting from the startColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        IXlsxRow AdjustToContents(Int32 startColumn);

        /// <summary>
        /// Adjusts the height of the row based on its contents, starting from the startColumn and ending at endColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        /// <param name="endColumn">The column to end calculating the row height.</param>
        IXlsxRow AdjustToContents(Int32 startColumn, Int32 endColumn);

        IXlsxRow AdjustToContents(Double minHeight, Double maxHeight);

        IXlsxRow AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight);

        /// <summary>
        /// Adjust height of the column according to the content of the cells.
        /// </summary>
        /// <param name="startColumn">Number of a first column whose content is considered.</param>
        /// <param name="endColumn">Number of a last column whose content is considered.</param>
        /// <param name="minHeightPt">Minimum height of adjusted column, in points.</param>
        /// <param name="maxHeightPt">Maximum height of adjusted column, in points.</param>
        IXlsxRow AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeightPt, Double maxHeightPt);

        /// <summary>Hides this row.</summary>
        IXlsxRow Hide();

        /// <summary>Unhides this row.</summary>
        IXlsxRow Unhide();

        /// <summary>
        /// Gets a value indicating whether this row is hidden or not.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this row is hidden; otherwise, <c>false</c>.
        /// </value>
        Boolean IsHidden { get; }

        /// <summary>
        /// Gets or sets the outline level of this row.
        /// </summary>
        /// <value>
        /// The outline level of this row.
        /// </value>
        Int32 OutlineLevel { get; set; }

        /// <summary>
        /// Adds this row to the next outline level (Increments the outline level for this row by 1).
        /// </summary>
        IXlsxRow Group();

        /// <summary>
        /// Adds this row to the next outline level (Increments the outline level for this row by 1).
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
        IXlsxRow Group(Boolean collapse);

        /// <summary>
        /// Sets outline level for this row.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXlsxRow Group(Int32 outlineLevel);

        /// <summary>
        /// Sets outline level for this row.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
        IXlsxRow Group(Int32 outlineLevel, Boolean collapse);

        /// <summary>
        /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
        /// </summary>
        IXlsxRow Ungroup();

        /// <summary>
        /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove this row from all outline levels.</param>
        IXlsxRow Ungroup(Boolean fromAll);

        /// <summary>
        /// Show this row as collapsed.
        /// </summary>
        IXlsxRow Collapse();

        /// <summary>
        /// Gets the cell in the specified column.
        /// </summary>
        /// <param name="columnNumber">The cell's column.</param>
        IXlsxCell Cell(Int32 columnNumber);

        /// <summary>
        /// Gets the cell in the specified column.
        /// </summary>
        /// <param name="columnLetter">The cell's column.</param>
        IXlsxCell Cell(String columnLetter);

        /// <summary>
        /// Returns the specified group of cells, separated by commas.
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
        /// </summary>
        /// <param name="cellsInRow">The row's cells to return.</param>
        new IXlsxCells Cells(String cellsInRow);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstColumn">The first column in the group of cells to return.</param>
        /// <param name="lastColumn">The last column in the group of cells to return.</param>
        IXlsxCells Cells(Int32 firstColumn, Int32 lastColumn);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstColumn">The first column in the group of cells to return.</param>
        /// <param name="lastColumn">The last column in the group of cells to return.</param>
        IXlsxCells Cells(String firstColumn, String lastColumn);

        /// <summary>Expands this row (if it's collapsed).</summary>
        IXlsxRow Expand();

        Int32 CellCount();

        IXlsxRangeRow CopyTo(IXlsxCell cell);

        IXlsxRangeRow CopyTo(IXlsxRangeBase range);

        IXlsxRow CopyTo(IXlsxRow row);

        IXlsxRow Sort();

        IXlsxRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXlsxRangeRow Row(Int32 start, Int32 end);

        IXlsxRangeRow Row(IXlsxCell start, IXlsxCell end);

        IXlsxRangeRows Rows(String columns);

        /// <summary>
        /// Adds a horizontal page break after this row.
        /// </summary>
        IXlsxRow AddHorizontalPageBreak();

        IXlsxRow RowAbove();

        IXlsxRow RowAbove(Int32 step);

        IXlsxRow RowBelow();

        IXlsxRow RowBelow(Int32 step);

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXlsxRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

        IXlsxRangeRow RowUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents);
    }
}
