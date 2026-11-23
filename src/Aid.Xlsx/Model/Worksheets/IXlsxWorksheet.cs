using Aid.Xlsx.Calculation.Exceptions;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Sparklines;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.Worksheets;

public enum XlsxWorksheetVisibility { Visible, Hidden, VeryHidden }

public interface IXlsxWorksheet : IXlsxRangeBase, IXlsxProtectable<IXlsxSheetProtection, XlsxSheetProtectionElements>
{
    /// <summary>
    /// Gets the workbook that contains this worksheet
    /// </summary>
    XlsxWorkbook Workbook { get; }

    /// <summary>
    /// Gets or sets the default column width for this worksheet.
    /// </summary>
    double ColumnWidth { get; set; }

    /// <summary>
    /// Gets or sets the default row height for this worksheet.
    /// </summary>
    double RowHeight { get; set; }

    /// <summary>
    /// Gets or sets the name (caption) of this worksheet. The sheet rename also renames sheet
    /// in formulas and defined names.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// Gets or sets the position of the sheet.
    /// <para>When setting the Position all other sheets' positions are shifted accordingly.</para>
    /// </summary>
    int Position { get; set; }

    /// <summary>
    /// Gets an object to manipulate the sheet's print options.
    /// </summary>
    IXlsxPageSetup PageSetup { get; }

    /// <summary>
    /// Gets an object to manipulate the Outline levels.
    /// </summary>
    IXlsxOutline Outline { get; }

    /// <summary>
    /// All hyperlinks in the sheet.
    /// </summary>
    IXlsxHyperlinks Hyperlinks { get; }

    /// <summary>
    /// Gets the first row of the worksheet.
    /// </summary>
    IXlsxRow FirstRow();

    /// <summary>
    /// Gets the first non-empty row of the worksheet that contains a cell with a value.
    /// <para>Formatted empty cells do not count.</para>
    /// </summary>
    IXlsxRow? FirstRowUsed();

    /// <summary>
    /// Gets the first non-empty row of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    IXlsxRow? FirstRowUsed(XlsxCellsUsedOptions options);

    /// <summary>
    /// Gets the last row of the worksheet.
    /// </summary>
    IXlsxRow LastRow();

    /// <summary>
    /// Gets the last non-empty row of the worksheet that contains a cell with a value.
    /// </summary>
    IXlsxRow? LastRowUsed();

    /// <summary>
    /// Gets the last non-empty row of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    IXlsxRow? LastRowUsed(XlsxCellsUsedOptions options);

    /// <summary>
    /// Gets the first column of the worksheet.
    /// </summary>
    IXlsxColumn FirstColumn();

    /// <summary>
    /// Gets the first non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    IXlsxColumn? FirstColumnUsed();

    /// <summary>
    /// Gets the first non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    IXlsxColumn? FirstColumnUsed(XlsxCellsUsedOptions options);

    /// <summary>
    /// Gets the last column of the worksheet.
    /// </summary>
    IXlsxColumn LastColumn();

    /// <summary>
    /// Gets the last non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    IXlsxColumn? LastColumnUsed();

    /// <summary>
    /// Gets the last non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    IXlsxColumn? LastColumnUsed(XlsxCellsUsedOptions options);

    /// <summary>
    /// Gets a collection of all columns in this worksheet.
    /// </summary>
    IXlsxColumns Columns();

    /// <summary>
    /// Gets a collection of the specified columns in this worksheet, separated by commas.
    /// <para>e.g. Columns("G:H"), Columns("10:11,13:14"), Columns("P:Q,S:T"), Columns("V")</para>
    /// </summary>
    /// <param name="columns">The columns to return.</param>
    IXlsxColumns Columns(string columns);

    /// <summary>
    /// Gets a collection of the specified columns in this worksheet.
    /// </summary>
    /// <param name="firstColumn">The first column to return.</param>
    /// <param name="lastColumn">The last column to return.</param>
    IXlsxColumns Columns(string firstColumn, string lastColumn);

    /// <summary>
    /// Gets a collection of the specified columns in this worksheet.
    /// </summary>
    /// <param name="firstColumn">The first column to return.</param>
    /// <param name="lastColumn">The last column to return.</param>
    IXlsxColumns Columns(int firstColumn, int lastColumn);

    /// <summary>
    /// Gets a collection of all rows in this worksheet.
    /// </summary>
    IXlsxRows Rows();

    /// <summary>
    /// Gets a collection of the specified rows in this worksheet, separated by commas.
    /// <para>e.g. Rows("4:5"), Rows("7:8,10:11"), Rows("13")</para>
    /// </summary>
    /// <param name="rows">The rows to return.</param>
    IXlsxRows Rows(string rows);

    /// <summary>
    /// Gets a collection of the specified rows in this worksheet.
    /// </summary>
    /// <param name="firstRow">The first row to return.</param>
    /// <param name="lastRow">The last row to return.</param>
    IXlsxRows Rows(int firstRow, int lastRow);

    /// <summary>
    /// Gets the specified row of the worksheet.
    /// </summary>
    /// <param name="row">The worksheet's row.</param>
    IXlsxRow Row(int row);

    /// <summary>
    /// Gets the specified column of the worksheet.
    /// </summary>
    /// <param name="column">The worksheet's column.</param>
    IXlsxColumn Column(int column);

    /// <summary>
    /// Gets the specified column of the worksheet.
    /// </summary>
    /// <param name="column">The worksheet's column.</param>
    IXlsxColumn Column(string column);

    /// <summary>
    /// Gets the cell at the specified row and column.
    /// </summary>
    /// <param name="row">The cell's row.</param>
    /// <param name="column">The cell's column.</param>
    IXlsxCell Cell(int row, int column);

    /// <summary>Gets the cell at the specified address.</summary>
    /// <param name="cellAddressInRange">The cell address in the worksheet.</param>
    /// <exception cref="ArgumentException">Address is not A1 or workbook-scoped named range.</exception>
    IXlsxCell Cell(string cellAddressInRange);

    /// <summary>
    /// Gets the cell at the specified row and column.
    /// </summary>
    /// <param name="row">The cell's row.</param>
    /// <param name="column">The cell's column.</param>
    IXlsxCell Cell(int row, string column);

    /// <summary>Gets the cell at the specified address.</summary>
    /// <param name="cellAddressInRange">The cell address in the worksheet.</param>
    IXlsxCell Cell(IXlsxAddress cellAddressInRange);

    /// <summary>
    /// Returns the specified range.
    /// </summary>
    /// <param name="rangeAddress">The range boundaries.</param>
    IXlsxRange Range(IXlsxRangeAddress rangeAddress);

    /// <summary>Returns the specified range.</summary>
    /// <para>e.g. Range("A1"), Range("A1:C2")</para>
    /// <param name="rangeAddress">The range boundaries.</param>
    /// <exception cref="ArgumentException"><paramref name="rangeAddress"/> is not a valid address or named range.</exception>
    IXlsxRange Range(string rangeAddress);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCell">The first cell in the range.</param>
    /// <param name="lastCell"> The last cell in the range.</param>
    IXlsxRange Range(IXlsxCell firstCell, IXlsxCell lastCell);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress"> The last cell address in the worksheet.</param>
    IXlsxRange Range(string firstCellAddress, string lastCellAddress);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress"> The last cell address in the worksheet.</param>
    IXlsxRange Range(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress);

    /// <summary>Returns a collection of ranges, separated by commas.</summary>
    /// <para>e.g. Ranges("A1"), Ranges("A1:C2"), Ranges("A1:B2,D1:D4")</para>
    /// <param name="ranges">The ranges to return.</param>
    IXlsxRanges Ranges(string ranges);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCellRow">   The first cell's row of the range to return.</param>
    /// <param name="firstCellColumn">The first cell's column of the range to return.</param>
    /// <param name="lastCellRow">    The last cell's row of the range to return.</param>
    /// <param name="lastCellColumn"> The last cell's column of the range to return.</param>
    /// <returns>.</returns>
    IXlsxRange Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

    /// <summary>Gets the number of rows in this worksheet.</summary>
    int RowCount();

    /// <summary>Gets the number of columns in this worksheet.</summary>
    int ColumnCount();

    /// <summary>
    /// Collapses all outlined rows.
    /// </summary>
    IXlsxWorksheet CollapseRows();

    /// <summary>
    /// Collapses all outlined columns.
    /// </summary>
    IXlsxWorksheet CollapseColumns();

    /// <summary>
    /// Expands all outlined rows.
    /// </summary>
    IXlsxWorksheet ExpandRows();

    /// <summary>
    /// Expands all outlined columns.
    /// </summary>
    IXlsxWorksheet ExpandColumns();

    /// <summary>
    /// Collapses the outlined rows of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    IXlsxWorksheet CollapseRows(int outlineLevel);

    /// <summary>
    /// Collapses the outlined columns of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    IXlsxWorksheet CollapseColumns(int outlineLevel);

    /// <summary>
    /// Expands the outlined rows of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    IXlsxWorksheet ExpandRows(int outlineLevel);

    /// <summary>
    /// Expands the outlined columns of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    IXlsxWorksheet ExpandColumns(int outlineLevel);

    /// <summary>
    /// Deletes this worksheet.
    /// </summary>
    void Delete();

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    IXlsxDefinedNames NamedRanges { get; }

    /// <summary>
    /// Gets an object to manage this worksheet's defined names.
    /// </summary>
    IXlsxDefinedNames DefinedNames { get; }

    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    IXlsxDefinedName NamedRange(string rangeName);

    /// <summary>
    /// Gets the specified defined name.
    /// </summary>
    /// <param name="name">Name identifier of defined name, without sheet name.</param>
    /// <exception cref="ArgumentException">Name wasn't found in sheets defined names.</exception>
    IXlsxDefinedName DefinedName(string name);

    /// <summary>
    /// Gets an object to manage how the worksheet is going to displayed by Excel.
    /// </summary>
    IXlsxSheetView SheetView { get; }

    /// <summary>
    /// Gets the Excel table of the given index
    /// </summary>
    /// <param name="index">Index of the table to return</param>
    IXlsxTable Table(int index);

    /// <summary>
    /// Gets the Excel table of the given name
    /// </summary>
    /// <param name="name">Name of the table to return</param>
    IXlsxTable Table(string name);

    /// <summary>
    /// Gets an object to manage this worksheet's Excel tables
    /// </summary>
    IXlsxTables Tables { get; }

    /// <summary>
    /// Copies the
    /// </summary>
    /// <param name="newSheetName"></param>
    IXlsxWorksheet CopyTo(string newSheetName);

    IXlsxWorksheet CopyTo(string newSheetName, int position);

    IXlsxWorksheet CopyTo(XlsxWorkbook workbook);

    /// <summary>
    /// Copy a worksheet from this workbook to a different workbook as a new sheet.
    /// </summary>
    /// <param name="workbook">Workbook into which copy this sheet.</param>
    /// <param name="newSheetName">Name of new sheet in the <paramref name="workbook"/> where will the data be copied. Sheet will be in the last position.</param>
    /// <returns>Newly created sheet in the <paramref name="workbook"/>.</returns>
    IXlsxWorksheet CopyTo(XlsxWorkbook workbook, string newSheetName);

    IXlsxWorksheet CopyTo(XlsxWorkbook workbook, string newSheetName, int position);

    IXlsxRange? RangeUsed();

    IXlsxRange? RangeUsed(XlsxCellsUsedOptions options);

    IXlsxDataValidations DataValidations { get; }

    XlsxWorksheetVisibility Visibility { get; set; }

    IXlsxWorksheet Hide();

    IXlsxWorksheet Unhide();

    IXlsxSortElements SortRows { get; }

    IXlsxSortElements SortColumns { get; }

    IXlsxRange Sort();

    IXlsxRange Sort(string columnsToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

    IXlsxRange Sort(int columnToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

    IXlsxRange SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

    //IXlsxCharts Charts { get; }

    bool ShowFormulas { get; set; }

    bool ShowGridLines { get; set; }

    bool ShowOutlineSymbols { get; set; }

    bool ShowRowColHeaders { get; set; }

    bool ShowRuler { get; set; }

    bool ShowWhiteSpace { get; set; }

    bool ShowZeros { get; set; }

    IXlsxWorksheet SetShowFormulas(); IXlsxWorksheet SetShowFormulas(bool value);

    IXlsxWorksheet SetShowGridLines(); IXlsxWorksheet SetShowGridLines(bool value);

    IXlsxWorksheet SetShowOutlineSymbols(); IXlsxWorksheet SetShowOutlineSymbols(bool value);

    IXlsxWorksheet SetShowRowColHeaders(); IXlsxWorksheet SetShowRowColHeaders(bool value);

    IXlsxWorksheet SetShowRuler(); IXlsxWorksheet SetShowRuler(bool value);

    IXlsxWorksheet SetShowWhiteSpace(); IXlsxWorksheet SetShowWhiteSpace(bool value);

    IXlsxWorksheet SetShowZeros(); IXlsxWorksheet SetShowZeros(bool value);

    XlsxColor TabColor { get; set; }

    IXlsxWorksheet SetTabColor(XlsxColor color);

    bool TabSelected { get; set; }

    bool TabActive { get; set; }

    IXlsxWorksheet SetTabSelected(); IXlsxWorksheet SetTabSelected(bool value);

    IXlsxWorksheet SetTabActive(); IXlsxWorksheet SetTabActive(bool value);

    IXlsxPivotTable PivotTable(string name);

    IXlsxPivotTables PivotTables { get; }

    bool RightToLeft { get; set; }

    IXlsxWorksheet SetRightToLeft(); IXlsxWorksheet SetRightToLeft(bool value);

    IXlsxAutoFilter AutoFilter { get; }

    IXlsxRows RowsUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents, Func<IXlsxRow, bool>? predicate = null);

    IXlsxRows RowsUsed(Func<IXlsxRow, bool>? predicate);

    IXlsxColumns ColumnsUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents, Func<IXlsxColumn, bool>? predicate = null);

    IXlsxColumns ColumnsUsed(Func<IXlsxColumn, bool>? predicate);

    IXlsxRanges MergedRanges { get; }

    IXlsxConditionalFormats ConditionalFormats { get; }

    IXlsxSparklineGroups SparklineGroups { get; }

    IXlsxRanges SelectedRanges { get; }

    /// <summary>
    /// The active cell of the worksheet.
    /// </summary>
    IXlsxCell? ActiveCell { get; set; }

    /// <summary>
    /// Evaluate an formula and return a result.
    /// </summary>
    /// <param name="expression">Formula to evaluate.</param>
    /// <param name="formulaAddress">A cell address that is used to provide context for formula calculation (mostly implicit intersection).</param>
    /// <exception cref="MissingContextException">If <paramref name="formulaAddress"/> was needed for some part of calculation.</exception>
    XlsxCellValue Evaluate(string expression, string? formulaAddress = null);

    /// <summary>
    /// Force recalculation of all cell formulas in the sheet while leaving other sheets without change, even if their dirty cells.
    /// </summary>
    void RecalculateAllFormulas();

    string Author { get; set; }

    IXlsxPictures Pictures { get; }

    IXlsxPicture Picture(string pictureName);

    IXlsxPicture AddPicture(Stream stream);

    IXlsxPicture AddPicture(Stream stream, string name);

    IXlsxPicture AddPicture(Stream stream, XlsxPictureFormat format);

    IXlsxPicture AddPicture(Stream stream, XlsxPictureFormat format, string name);

    IXlsxPicture AddPicture(string imageFile);

    IXlsxPicture AddPicture(string imageFile, string name);
}
