#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Tables;

namespace Aid.Xlsx.Model.Grid.Ranges;

public interface IXlsxRangeColumn : IXlsxRangeBase
{
    /// <summary>
    /// Gets the cell in the specified row.
    /// </summary>
    /// <param name="rowNumber">The cell's row.</param>
    IXlsxCell Cell(Int32 rowNumber);

    /// <summary>
    /// Returns the specified group of cells, separated by commas.
    /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1:2,4:5")</para>
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
    /// Inserts X number of columns to the right of this range.
    /// <para>All cells to the right of this range will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of columns to insert.</param>
    IXlsxRangeColumns InsertColumnsAfter(int numberOfColumns);

    IXlsxRangeColumns InsertColumnsAfter(int numberOfColumns, Boolean expandRange);

    /// <summary>
    /// Inserts X number of columns to the left of this range.
    /// <para>This range and all cells to the right of this range will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of columns to insert.</param>
    IXlsxRangeColumns InsertColumnsBefore(int numberOfColumns);

    IXlsxRangeColumns InsertColumnsBefore(int numberOfColumns, Boolean expandRange);

    /// <summary>
    /// Inserts X number of cells on top of this column.
    /// <para>This column and all cells below it will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of cells to insert.</param>
    IXlsxCells InsertCellsAbove(int numberOfRows);

    IXlsxCells InsertCellsAbove(int numberOfRows, Boolean expandRange);

    /// <summary>
    /// Inserts X number of cells below this range.
    /// <para>All cells below this column will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of cells to insert.</param>
    IXlsxCells InsertCellsBelow(int numberOfRows);

    IXlsxCells InsertCellsBelow(int numberOfRows, Boolean expandRange);

    /// <summary>
    /// Deletes this range and shifts the cells at the right.
    /// </summary>
    void Delete();

    /// <summary>
    /// Deletes this range and shifts the surrounding cells accordingly.
    /// </summary>
    /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
    void Delete(XlsxShiftDeletedCells shiftDeleteCells);

    /// <summary>
    /// Gets this column's number in the range
    /// </summary>
    Int32 ColumnNumber();

    /// <summary>
    /// Gets this column's letter in the range
    /// </summary>
    String ColumnLetter();

    Int32 CellCount();

    IXlsxRangeColumn CopyTo(IXlsxCell target);

    IXlsxRangeColumn CopyTo(IXlsxRangeBase target);

    IXlsxRangeColumn Sort(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

    IXlsxRangeColumn Column(Int32 start, Int32 end);

    IXlsxRangeColumn Column(IXlsxCell start, IXlsxCell end);

    IXlsxRangeColumns Columns(String columns);

    IXlsxRangeColumn ColumnLeft();

    IXlsxRangeColumn ColumnLeft(Int32 step);

    IXlsxRangeColumn ColumnRight();

    IXlsxRangeColumn ColumnRight(Int32 step);

    IXlsxColumn WorksheetColumn();

    IXlsxTable AsTable();

    IXlsxTable AsTable(String name);

    IXlsxTable CreateTable();

    IXlsxTable CreateTable(String name);

    /// <summary>
    /// Clears the contents of this column.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    new IXlsxRangeColumn Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    IXlsxRangeColumn ColumnUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents);
}
