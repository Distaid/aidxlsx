#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Data;
using System.Globalization;
using Aid.Xlsx.Calculation.Exceptions;
using Aid.Xlsx.Model.CustomProperties;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Workbooks;

public interface IXlsxWorkbook : IXlsxProtectable<IXlsxWorkbookProtection, XlsxWorkbookProtectionElements>, IDisposable
{
    string Author { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's calculation mode.
    /// </summary>
    XlsxCalculateMode CalculateMode { get; set; }

    bool CalculationOnSave { get; set; }

    /// <summary>
    ///   Gets or sets the default column width for the workbook.
    ///   <para>All new worksheets will use this column width.</para>
    /// </summary>
    double ColumnWidth { get; set; }

    IXlsxCustomProperties CustomProperties { get; }

    bool DefaultRightToLeft { get; }

    bool DefaultShowFormulas { get; }

    bool DefaultShowGridLines { get; }

    bool DefaultShowOutlineSymbols { get; }

    bool DefaultShowRowColHeaders { get; }

    bool DefaultShowRuler { get; }

    bool DefaultShowWhiteSpace { get; }

    bool DefaultShowZeros { get; }

    IXlsxFileSharing FileSharing { get; }

    bool ForceFullCalculation { get; set; }

    bool FullCalculationOnLoad { get; set; }

    bool FullPrecision { get; set; }

    bool LockStructure { get; set; }

    bool LockWindows { get; set; }

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    IXlsxDefinedNames NamedRanges { get; }

    /// <summary>
    ///   Gets an object to manipulate this workbook's defined names.
    /// </summary>
    IXlsxDefinedNames DefinedNames { get; }

    /// <summary>
    ///   Gets or sets the default outline options for the workbook.
    ///   <para>All new worksheets will use these outline options.</para>
    /// </summary>
    IXlsxOutline Outline { get; set; }

    /// <summary>
    ///   Gets or sets the default page options for the workbook.
    ///   <para>All new worksheets will use these page options.</para>
    /// </summary>
    IXlsxPageSetup PageOptions { get; set; }

    /// <summary>
    ///   Gets all pivot caches in a workbook. A one cache can be
    ///   used by multiple tables. Unused caches are not saved.
    /// </summary>
    IXlsxPivotCaches PivotCaches { get; }

    /// <summary>
    ///   Gets or sets the workbook's properties.
    /// </summary>
    XlsxWorkbookProperties Properties { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's reference style.
    /// </summary>
    XlsxReferenceStyle ReferenceStyle { get; set; }

    bool RightToLeft { get; set; }

    /// <summary>
    ///   Gets or sets the default row height for the workbook.
    ///   <para>All new worksheets will use this row height.</para>
    /// </summary>
    double RowHeight { get; set; }

    bool ShowFormulas { get; set; }

    bool ShowGridLines { get; set; }

    bool ShowOutlineSymbols { get; set; }

    bool ShowRowColHeaders { get; set; }

    bool ShowRuler { get; set; }

    bool ShowWhiteSpace { get; set; }

    bool ShowZeros { get; set; }

    /// <summary>
    /// Gets or sets the default format of the workbook. All cells that don't have specified format, either
    /// at sheet, column, row or cell level, will use this format.
    /// </summary>
    IXlsxStyle Style { get; set; }

    /// <summary>
    ///   Gets an object to manipulate this workbook's theme.
    /// </summary>
    IXlsxTheme Theme { get; }

    bool Use1904DateSystem { get; set; }

    /// <summary>
    ///   Gets an object to manipulate the worksheets.
    /// </summary>
    IXlsxWorksheets Worksheets { get; }

    IXlsxWorksheet AddWorksheet();

    IXlsxWorksheet AddWorksheet(int position);

    IXlsxWorksheet AddWorksheet(string sheetName);

    IXlsxWorksheet AddWorksheet(string sheetName, int position);

    void AddWorksheet(DataSet dataSet);

    void AddWorksheet(IXlsxWorksheet worksheet);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1). The dataTable's name is used for the
    /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
    /// </summary>
    /// <param name="dataTable">Datatable to insert</param>
    /// <returns>Inserted Worksheet</returns>
    IXlsxWorksheet AddWorksheet(DataTable dataTable);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1). The sheetName provided is used for the
    /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
    /// </summary>
    /// <param name="dataTable">dataTable to insert as Excel Table</param>
    /// <param name="sheetName">Worksheet and Excel Table name</param>
    /// <returns>Inserted Worksheet</returns>
    IXlsxWorksheet AddWorksheet(DataTable dataTable, string sheetName);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1).
    /// </summary>
    /// <param name="dataTable">dataTable to insert as Excel Table</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="tableName">Excel Table name</param>
    /// <returns>Inserted Worksheet</returns>
    IXlsxWorksheet AddWorksheet(DataTable dataTable, string sheetName, string tableName);

    IXlsxCell Cell(string namedCell);

    IXlsxCells Cells(string namedCells);

    IXlsxCustomProperty CustomProperty(string name);

    /// <summary>
    /// Evaluate a formula expression.
    /// </summary>
    /// <param name="expression">Formula expression to evaluate.</param>
    /// <exception cref="MissingContextException">
    /// If the expression contains a function that requires a context (e.g. current cell or worksheet).
    /// </exception>
    XlsxCellValue Evaluate(string expression);

    IXlsxCells FindCells(Func<IXlsxCell, bool> predicate);

    IXlsxColumns FindColumns(Func<IXlsxColumn, bool> predicate);

    IXlsxRows FindRows(Func<IXlsxRow, bool> predicate);

#nullable enable
    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    IXlsxDefinedName? NamedRange(string name);

    /// <summary>
    /// Try to find a defined name. If <paramref name="name"/> specifies a sheet, try to find
    /// name in the sheet first and fall back to the workbook if not found in the sheet.
    /// <para>
    /// <example>
    /// Requested name <c>Sheet1!Name</c> will first try to find <c>Name</c> in a sheet
    /// <c>Sheet1</c> (if such sheet exists) and if not found there, tries to find <c>Name</c>
    /// in workbook.
    /// </example>
    /// </para>
    /// <para>
    /// <example>
    /// Requested name <c>Name</c> will be searched only in a workbooks <see cref="DefinedNames"/>.
    /// </example>
    /// </para>
    /// </summary>
    /// <param name="name">Name of requested name, either plain name (e.g. <c>Name</c>) or with
    /// sheet specified (e.g. <c>Sheet!Name</c>).</param>
    /// <returns>Found name or null.</returns>
    IXlsxDefinedName? DefinedName(string name);
#nullable disable

    IXlsxRange Range(string range);

    IXlsxRange RangeFromFullAddress(string rangeAddress, out IXlsxWorksheet ws);

    IXlsxRanges Ranges(string ranges);

    /// <summary>
    /// Force recalculation of all cell formulas.
    /// </summary>
    void RecalculateAllFormulas();

    /// <summary>
    ///   Saves the current workbook.
    /// </summary>
    void Save();

    /// <summary>
    ///   Saves the current workbook and optionally performs validation
    /// </summary>
    void Save(bool validate, bool evaluateFormulae = false);

    void Save(SaveOptions options);

    /// <summary>
    ///   Saves the current workbook to a file.
    /// </summary>
    void SaveAs(string file);

    /// <summary>
    ///   Saves the current workbook to a file and optionally validates it.
    /// </summary>
    void SaveAs(string file, bool validate, bool evaluateFormulae = false);

    void SaveAs(string file, SaveOptions options);

    /// <summary>
    ///   Saves the current workbook to a stream.
    /// </summary>
    void SaveAs(Stream stream);

    /// <summary>
    ///   Saves the current workbook to a stream and optionally validates it.
    /// </summary>
    void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false);

    void SaveAs(Stream stream, SaveOptions options);

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    IEnumerable<IXlsxCell> Search(string searchText, CompareOptions compareOptions = CompareOptions.Ordinal, bool searchFormulae = false);

    XlsxWorkbook SetLockStructure(bool value);

    XlsxWorkbook SetLockWindows(bool value);

    XlsxWorkbook SetUse1904DateSystem();

    XlsxWorkbook SetUse1904DateSystem(bool value);

    /// <summary>
    /// Gets the Excel table of the given name
    /// </summary>
    /// <param name="tableName">Name of the table to return.</param>
    /// <param name="comparisonType">One of the enumeration values that specifies how the strings will be compared.</param>
    /// <returns>The table with given name</returns>
    /// <exception cref="ArgumentOutOfRangeException">If no tables with this name could be found in the workbook.</exception>
    IXlsxTable Table(string tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase);

    bool TryGetWorksheet(string name, out IXlsxWorksheet worksheet);

    IXlsxWorksheet Worksheet(string name);

    IXlsxWorksheet Worksheet(int position);
}
