#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;

namespace Aid.Xlsx.Model.Grid.Rows;

public interface IXlsxRow : IXlsxRangeBase
{
    /// <summary>
    /// Gets or sets the height of this row.
    /// </summary>
    /// <value>
    /// The width of this row in points.
    /// </value>
    double Height { get; set; }

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
    int RowNumber();

    /// <summary>
    /// Inserts X number of rows below this one.
    /// <para>All rows below will be shifted accordingly.</para>
    /// </summary>
    /// <param name="numberOfRows">The number of rows to insert.</param>
    IXlsxRows InsertRowsBelow(int numberOfRows);

    /// <summary>
    /// Inserts X number of rows above this one.
    /// <para>This row and all below will be shifted accordingly.</para>
    /// </summary>
    /// <param name="numberOfRows">The number of rows to insert.</param>
    IXlsxRows InsertRowsAbove(int numberOfRows);

    IXlsxRow AdjustToContents();

    /// <summary>
    /// Adjusts the height of the row based on its contents, starting from the startColumn.
    /// </summary>
    /// <param name="startColumn">The column to start calculating the row height.</param>
    IXlsxRow AdjustToContents(int startColumn);

    /// <summary>
    /// Adjusts the height of the row based on its contents, starting from the startColumn and ending at endColumn.
    /// </summary>
    /// <param name="startColumn">The column to start calculating the row height.</param>
    /// <param name="endColumn">The column to end calculating the row height.</param>
    IXlsxRow AdjustToContents(int startColumn, int endColumn);

    IXlsxRow AdjustToContents(double minHeight, double maxHeight);

    IXlsxRow AdjustToContents(int startColumn, double minHeight, double maxHeight);

    /// <summary>
    /// Adjust height of the column according to the content of the cells.
    /// </summary>
    /// <param name="startColumn">Number of a first column whose content is considered.</param>
    /// <param name="endColumn">Number of a last column whose content is considered.</param>
    /// <param name="minHeightPt">Minimum height of adjusted column, in points.</param>
    /// <param name="maxHeightPt">Maximum height of adjusted column, in points.</param>
    IXlsxRow AdjustToContents(int startColumn, int endColumn, double minHeightPt, double maxHeightPt);

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
    bool IsHidden { get; }

    /// <summary>
    /// Gets or sets the outline level of this row.
    /// </summary>
    /// <value>
    /// The outline level of this row.
    /// </value>
    int OutlineLevel { get; set; }

    /// <summary>
    /// Adds this row to the next outline level (Increments the outline level for this row by 1).
    /// </summary>
    IXlsxRow Group();

    /// <summary>
    /// Adds this row to the next outline level (Increments the outline level for this row by 1).
    /// </summary>
    /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
    IXlsxRow Group(bool collapse);

    /// <summary>
    /// Sets outline level for this row.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    IXlsxRow Group(int outlineLevel);

    /// <summary>
    /// Sets outline level for this row.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
    IXlsxRow Group(int outlineLevel, bool collapse);

    /// <summary>
    /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
    /// </summary>
    IXlsxRow Ungroup();

    /// <summary>
    /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
    /// </summary>
    /// <param name="fromAll">If set to <c>true</c> it will remove this row from all outline levels.</param>
    IXlsxRow Ungroup(bool fromAll);

    /// <summary>
    /// Show this row as collapsed.
    /// </summary>
    IXlsxRow Collapse();

    /// <summary>
    /// Gets the cell in the specified column.
    /// </summary>
    /// <param name="columnNumber">The cell's column.</param>
    IXlsxCell Cell(int columnNumber);

    /// <summary>
    /// Gets the cell in the specified column.
    /// </summary>
    /// <param name="columnLetter">The cell's column.</param>
    IXlsxCell Cell(string columnLetter);

    /// <summary>
    /// Returns the specified group of cells, separated by commas.
    /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
    /// </summary>
    /// <param name="cellsInRow">The row's cells to return.</param>
    new IXlsxCells Cells(string cellsInRow);

    /// <summary>
    /// Returns the specified group of cells.
    /// </summary>
    /// <param name="firstColumn">The first column in the group of cells to return.</param>
    /// <param name="lastColumn">The last column in the group of cells to return.</param>
    IXlsxCells Cells(int firstColumn, int lastColumn);

    /// <summary>
    /// Returns the specified group of cells.
    /// </summary>
    /// <param name="firstColumn">The first column in the group of cells to return.</param>
    /// <param name="lastColumn">The last column in the group of cells to return.</param>
    IXlsxCells Cells(string firstColumn, string lastColumn);

    /// <summary>Expands this row (if it's collapsed).</summary>
    IXlsxRow Expand();

    int CellCount();

    IXlsxRangeRow CopyTo(IXlsxCell cell);

    IXlsxRangeRow CopyTo(IXlsxRangeBase range);

    IXlsxRow CopyTo(IXlsxRow row);

    IXlsxRow Sort();

    IXlsxRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

    IXlsxRangeRow Row(int start, int end);

    IXlsxRangeRow Row(IXlsxCell start, IXlsxCell end);

    IXlsxRangeRows Rows(string columns);

    /// <summary>
    /// Adds a horizontal page break after this row.
    /// </summary>
    IXlsxRow AddHorizontalPageBreak();

    IXlsxRow RowAbove();

    IXlsxRow RowAbove(int step);

    IXlsxRow RowBelow();

    IXlsxRow RowBelow(int step);

    /// <summary>
    /// Clears the contents of this row.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    new IXlsxRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    IXlsxRangeRow RowUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents);
}
