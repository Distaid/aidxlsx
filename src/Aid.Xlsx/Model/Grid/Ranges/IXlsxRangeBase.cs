#nullable disable

using System.Globalization;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges;

public enum XlsxScope
{
    Workbook,
    Worksheet
}

public interface IXlsxRangeBase : IXlsxAddressable
{
    IXlsxWorksheet Worksheet { get; }

    /// <summary>
    /// Sets a value to every cell in this range.
    /// <para>
    /// Setter will clear a formula, if the cell contains a formula.
    /// If the value is a text that starts with a single quote, setter will prefix the value with a single quote through
    /// <see cref="IXlsxStyle.IncludeQuotePrefix"/> in Excel too and the value of cell is set to to non-quoted text.
    /// </para>
    /// </summary>
    XlsxCellValue Value { set; }

    /// <summary>
    ///   Sets the cells' formula with A1 references.
    /// </summary>
    /// <remarks>
    /// Setter trims the formula and if formula starts with an <c>=</c>, it is removed. If the
    /// formula contains unprefixed future function (e.g. <c>CONCAT</c>), it will be correctly
    /// prefixed (e.g. <c>_xlfn.CONCAT</c>).
    /// </remarks>
    /// <value>The formula with A1 references.</value>
    String FormulaA1 { set; }

    /// <summary>
    /// Create an array formula for all cells in the range.
    /// </summary>
    /// <remarks>
    /// Setter trims the formula and if formula starts with an <c>=</c>, it is removed. If the
    /// formula contains unprefixed future function (e.g. <c>CONCAT</c>), it will be correctly
    /// prefixed (e.g. <c>_xlfn.CONCAT</c>).
    /// </remarks>
    /// <exception cref="InvalidOperationException">When the range overlaps with a table, pivot table, merged cells or partially overlaps another array formula.</exception>
    String FormulaArrayA1 { set; }

    /// <summary>
    ///   Sets the cells' formula with R1C1 references.
    /// </summary>
    /// <remarks>
    /// Setter trims the formula and if formula starts with an <c>=</c>, it is removed. If the
    /// formula contains unprefixed future function (e.g. <c>CONCAT</c>), it will be correctly
    /// prefixed (e.g. <c>_xlfn.CONCAT</c>).
    /// </remarks>
    /// <value>The formula with R1C1 references.</value>
    String FormulaR1C1 { set; }

    IXlsxStyle Style { get; set; }

    /// <summary>
    ///   Gets or sets a value indicating whether this cell's text should be shared or not.
    /// </summary>
    /// <value>
    ///   If false the cell's text will not be shared and stored as an inline value.
    /// </value>
    Boolean ShareString { set; }

    /// <summary>
    ///   Returns the collection of cells.
    /// </summary>
    IXlsxCells Cells();

    IXlsxCells Cells(Boolean usedCellsOnly);

    IXlsxCells Cells(Boolean usedCellsOnly, XlsxCellsUsedOptions options);

    IXlsxCells Cells(String cells);

    IXlsxCells Cells(Func<IXlsxCell, Boolean> predicate);

    /// <summary>
    ///   Returns the collection of cells that have a value. Formats are ignored.
    /// </summary>
    IXlsxCells CellsUsed();

    /// <summary>
    /// Returns the collection of cells that have a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    IXlsxCells CellsUsed(XlsxCellsUsedOptions options);

    IXlsxCells CellsUsed(Func<IXlsxCell, Boolean> predicate);

    IXlsxCells CellsUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate);

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    IXlsxCells Search(String searchText, CompareOptions compareOptions = CompareOptions.Ordinal, Boolean searchFormulae = false);

    /// <summary>
    ///   Returns the first cell of this range.
    /// </summary>
    IXlsxCell FirstCell();

    /// <summary>
    ///   Returns the first non-empty cell with a value of this range. Formats are ignored.
    ///   <para>The cell's address is going to be ([First Row with a value], [First Column with a value])</para>
    /// </summary>
    IXlsxCell FirstCellUsed();

    /// <summary>
    /// Returns the first non-empty cell with a value of this range.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    IXlsxCell FirstCellUsed(XlsxCellsUsedOptions options);

    IXlsxCell FirstCellUsed(Func<IXlsxCell, Boolean> predicate);

    /// <summary>
    /// Returns the first non-empty cell with a value of this range.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    /// <param name="predicate">The predicate used to choose cells</param>
    IXlsxCell FirstCellUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate);

    /// <summary>
    ///   Returns the last cell of this range.
    /// </summary>
    IXlsxCell LastCell();

    /// <summary>
    ///   Returns the last non-empty cell with a value of this range. Formats are ignored.
    ///   <para>The cell's address is going to be ([Last Row with a value], [Last Column with a value])</para>
    /// </summary>
    IXlsxCell LastCellUsed();

    /// <summary>
    /// Returns the last non-empty cell with a value of this range.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    IXlsxCell LastCellUsed(XlsxCellsUsedOptions options);

    IXlsxCell LastCellUsed(Func<IXlsxCell, Boolean> predicate);

    IXlsxCell LastCellUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate);

    /// <summary>
    ///   Determines whether this range contains the specified range (completely).
    ///   <para>For partial matches use the range.Intersects method.</para>
    /// </summary>
    /// <param name = "rangeAddress">The range address.</param>
    /// <returns>
    ///   <c>true</c> if this range contains the specified range; otherwise, <c>false</c>.
    /// </returns>
    Boolean Contains(String rangeAddress);

    /// <summary>
    ///   Determines whether this range contains the specified range (completely).
    ///   <para>For partial matches use the range.Intersects method.</para>
    /// </summary>
    /// <param name = "range">The range to match.</param>
    /// <returns>
    ///   <c>true</c> if this range contains the specified range; otherwise, <c>false</c>.
    /// </returns>
    Boolean Contains(IXlsxRangeBase range);

    Boolean Contains(IXlsxCell cell);

    /// <summary>
    ///   Determines whether this range intersects the specified range.
    ///   <para>For whole matches use the range.Contains method.</para>
    /// </summary>
    /// <param name = "rangeAddress">The range address.</param>
    /// <returns>
    ///   <c>true</c> if this range intersects the specified range; otherwise, <c>false</c>.
    /// </returns>
    Boolean Intersects(String rangeAddress);

    /// <summary>
    ///   Determines whether this range contains the specified range.
    ///   <para>For whole matches use the range.Contains method.</para>
    /// </summary>
    /// <param name = "range">The range to match.</param>
    /// <returns>
    ///   <c>true</c> if this range intersects the specified range; otherwise, <c>false</c>.
    /// </returns>
    Boolean Intersects(IXlsxRangeBase range);

    /// <summary>
    ///   Unmerges this range.
    /// </summary>
    IXlsxRange Unmerge();

    /// <summary>
    /// Merges this range. Only the top-left cell will have a value, other values will be blank.
    /// </summary>
    IXlsxRange Merge();

    IXlsxRange Merge(Boolean checkIntersect);

    /// <summary>
    /// Creates/adds this range to workbook scoped <see cref="IXlsxDefinedNames"/>.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// </summary>
    /// <param name = "name">Name of the defined name, without sheet.</param>
    IXlsxRange AddToNamed(String name);

    /// <summary>
    /// Creates/adds this range to <see cref="IXlsxDefinedNames"/>.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// <param name = "name">Name of the defined name, without sheet.</param>
    /// <param name = "scope">The scope for the named range.</param>
    /// </summary>
    IXlsxRange AddToNamed(String name, XlsxScope scope);

    /// <summary>
    /// Creates/adds this range to <see cref="IXlsxDefinedNames"/>.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// <param name = "name">Name of the defined name, without sheet.</param>
    /// <param name = "scope">The scope for the named range.</param>
    /// <param name = "comment">The comments for the named range.</param>
    /// </summary>
    IXlsxRange AddToNamed(String name, XlsxScope scope, String comment);

    /// <summary>
    /// Clears the contents of this range.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    IXlsxRangeBase Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    /// <summary>
    ///   Deletes the cell comments from this range.
    /// </summary>
    void DeleteComments();

    /// <summary>
    /// Set value to all cells in the range.
    /// </summary>
    IXlsxRangeBase SetValue(XlsxCellValue value);

    /// <summary>
    ///   Converts this object to a range.
    /// </summary>
    IXlsxRange AsRange();

    Boolean IsMerged();

    Boolean IsEmpty();

    Boolean IsEmpty(XlsxCellsUsedOptions options);

    /// <summary>
    /// Determines whether range address spans the entire column.
    /// </summary>
    /// <returns>
    ///   <c>true</c> if is entire column; otherwise, <c>false</c>.
    /// </returns>
    Boolean IsEntireColumn();

    /// <summary>
    /// Determines whether range address spans the entire row.
    /// </summary>
    /// <returns>
    ///   <c>true</c> if is entire row; otherwise, <c>false</c>.
    /// </returns>

    Boolean IsEntireRow();

    /// <summary>
    /// Determines whether the range address spans the entire worksheet.
    /// </summary>
    /// <returns>
    ///   <c>true</c> if is entire sheet; otherwise, <c>false</c>.
    /// </returns>
    Boolean IsEntireSheet();

    IXlsxPivotTable CreatePivotTable(IXlsxCell targetCell, String name);

    //IXlsxChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn);

    IXlsxAutoFilter SetAutoFilter();

    IXlsxAutoFilter SetAutoFilter(Boolean value);

    /// <summary>
    /// Returns a data validation rule assigned to the range, if any, or creates a new instance of data validation rule if no rule exists.
    /// </summary>
    IXlsxDataValidation GetDataValidation();

    /// <summary>
    /// Creates a new data validation rule for the range, replacing the existing one.
    /// </summary>
    IXlsxDataValidation CreateDataValidation();

    IXlsxConditionalFormat AddConditionalFormat();

    void Select();

    /// <summary>
    /// Grows this the current range by one cell to each side
    /// </summary>
    IXlsxRangeBase Grow();

    /// <summary>
    /// Grows this the current range by the specified number of cells to each side.
    /// </summary>
    /// <param name="growCount">The grow count.</param>
    IXlsxRangeBase Grow(Int32 growCount);

    /// <summary>
    /// Shrinks this current range by one cell.
    /// </summary>
    IXlsxRangeBase Shrink();

    /// <summary>
    /// Shrinks the current range by the specified number of cells from each side.
    /// </summary>
    /// <param name="shrinkCount">The shrink count.</param>
    IXlsxRangeBase Shrink(Int32 shrinkCount);

    /// <summary>
    /// Returns the intersection of this range with another range on the same worksheet.
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
    /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
    /// <returns>The range address of the intersection</returns>
    IXlsxRangeAddress Intersection(IXlsxRangeBase otherRange, Func<IXlsxCell, Boolean> thisRangePredicate = null, Func<IXlsxCell, Boolean> otherRangePredicate = null);

    /// <summary>
    /// Returns the set of cells surrounding the current range.
    /// </summary>
    /// <param name="predicate">The predicate to apply on the resulting set of cells.</param>
    IXlsxCells SurroundingCells(Func<IXlsxCell, Boolean> predicate = null);

    /// <summary>
    /// Calculates the union of two ranges on the same worksheet.
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
    /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
    /// <returns>
    /// The union
    /// </returns>
    IXlsxCells Union(IXlsxRangeBase otherRange, Func<IXlsxCell, Boolean> thisRangePredicate = null, Func<IXlsxCell, Boolean> otherRangePredicate = null);

    /// <summary>
    /// Returns all cells in the current range that are not in the other range.
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
    /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
    IXlsxCells Difference(IXlsxRangeBase otherRange, Func<IXlsxCell, Boolean> thisRangePredicate = null, Func<IXlsxCell, Boolean> otherRangePredicate = null);

    /// <summary>
    /// Returns a range so that its offset from the target base range is equal to the offset of the current range to the source base range.
    /// For example, if the current range is D4:E4, the source base range is A1:C3, then the relative range to the target base range B10:D13 is E14:F14
    /// </summary>
    /// <param name="sourceBaseRange">The source base range.</param>
    /// <param name="targetBaseRange">The target base range.</param>
    /// <returns>The relative range</returns>
    IXlsxRangeBase Relative(IXlsxRangeBase sourceBaseRange, IXlsxRangeBase targetBaseRange);
}
