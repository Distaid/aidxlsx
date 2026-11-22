#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Grid.Rows;

namespace Aid.Xlsx.Model.Grid.Ranges;

public interface IXlsxRangeRow : IXlsxRangeBase
{
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
    /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1:2,4:5")</para>
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

    /// <summary>
    /// Inserts X number of cells to the right of this row.
    /// <para>All cells to the right of this row will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of cells to insert.</param>
    IXlsxCells InsertCellsAfter(int numberOfColumns);

    IXlsxCells InsertCellsAfter(int numberOfColumns, Boolean expandRange);

    /// <summary>
    /// Inserts X number of cells to the left of this row.
    /// <para>This row and all cells to the right of it will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of cells to insert.</param>
    IXlsxCells InsertCellsBefore(int numberOfColumns);

    IXlsxCells InsertCellsBefore(int numberOfColumns, Boolean expandRange);

    /// <summary>
    /// Inserts X number of rows on top of this row.
    /// <para>This row and all cells below it will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of rows to insert.</param>
    IXlsxRangeRows InsertRowsAbove(int numberOfRows);

    IXlsxRangeRows InsertRowsAbove(int numberOfRows, Boolean expandRange);

    /// <summary>
    /// Inserts X number of rows below this row.
    /// <para>All cells below this row will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of rows to insert.</param>
    IXlsxRangeRows InsertRowsBelow(int numberOfRows);

    IXlsxRangeRows InsertRowsBelow(int numberOfRows, Boolean expandRange);

    /// <summary>
    /// Deletes this range and shifts the cells below.
    /// </summary>
    void Delete();

    /// <summary>
    /// Deletes this range and shifts the surrounding cells accordingly.
    /// </summary>
    /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
    void Delete(XlsxShiftDeletedCells shiftDeleteCells);

    /// <summary>
    /// Gets this row's number in the range
    /// </summary>
    Int32 RowNumber();

    Int32 CellCount();

    IXlsxRangeRow CopyTo(IXlsxCell target);

    IXlsxRangeRow CopyTo(IXlsxRangeBase target);

    IXlsxRangeRow Sort();

    IXlsxRangeRow SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

    IXlsxRangeRow Row(Int32 start, Int32 end);

    IXlsxRangeRow Row(IXlsxCell start, IXlsxCell end);

    IXlsxRangeRows Rows(String rows);

    IXlsxRangeRow RowAbove();

    IXlsxRangeRow RowAbove(Int32 step);

    IXlsxRangeRow RowBelow();

    IXlsxRangeRow RowBelow(Int32 step);

    IXlsxRow WorksheetRow();

    /// <summary>
    /// Clears the contents of this row.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    new IXlsxRangeRow Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    IXlsxRangeRow RowUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents);
}
