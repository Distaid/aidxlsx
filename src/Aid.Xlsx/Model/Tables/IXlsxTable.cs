#nullable disable

using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Tables;

public interface IXlsxTable : IXlsxRange
{
    IXlsxAutoFilter AutoFilter { get; }
    IXlsxTableRange DataRange { get; }
    bool EmphasizeFirstColumn { get; set; }
    bool EmphasizeLastColumn { get; set; }
    IEnumerable<IXlsxTableField> Fields { get; }

    /// <summary>
    /// Change the name of a table. Structural references to the table are not updated.
    /// </summary>
    /// <exception cref="ArgumentException">If the new table name is already used by other table in the sheet.</exception>
    string Name { get; set; }
    bool ShowAutoFilter { get; set; }
    bool ShowColumnStripes { get; set; }
    bool ShowHeaderRow { get; set; }
    bool ShowRowStripes { get; set; }
    bool ShowTotalsRow { get; set; }
    XlsxTableTheme Theme { get; set; }

    /// <summary>
    /// Clears the contents of this table.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    new IXlsxTable Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    /// <summary>
    /// Get field of the table.
    /// </summary>
    /// <param name="fieldName">Name of the field. Field names are case-insensitive.</param>
    /// <returns>Requested field.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Table doesn't contain <paramref name="fieldName"/> field.</exception>
    IXlsxTableField Field(string fieldName);

    IXlsxTableField Field(int fieldIndex);

    IXlsxRangeRow HeadersRow();

    /// <summary>
    /// Appends the IEnumerable data elements and returns the range of the new rows.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    IXlsxRange AppendData(IEnumerable data, bool propagateExtraColumns = false);

    /// <summary>
    /// Appends the IEnumerable data elements and returns the range of the new rows.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    IXlsxRange AppendData(IEnumerable data, bool transpose, bool propagateExtraColumns = false);

    /// <summary>
    /// Appends the data of a data table and returns the range of the new rows.
    /// </summary>
    /// <param name="dataTable">The data table.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    IXlsxRange AppendData(DataTable dataTable, bool propagateExtraColumns = false);

    /// <summary>
    /// Appends the IEnumerable data elements and returns the range of the new rows.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="data">The table data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    IXlsxRange AppendData<T>(IEnumerable<T> data, bool propagateExtraColumns = false);

    /// <summary>
    /// Replaces the IEnumerable data elements and returns the table's data range.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    IXlsxRange ReplaceData(IEnumerable data, bool propagateExtraColumns = false);

    /// <summary>
    /// Replaces the IEnumerable data elements and returns the table's data range.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    IXlsxRange ReplaceData(IEnumerable data, bool transpose, bool propagateExtraColumns = false);

    /// <summary>
    /// Replaces the data from the records of a data table and returns the table's data range.
    /// </summary>
    /// <param name="dataTable">The data table.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    IXlsxRange ReplaceData(DataTable dataTable, bool propagateExtraColumns = false);

    /// <summary>
    /// Replaces the IEnumerable data elements as a table and the table's data range.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="data">The table data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    IXlsxRange ReplaceData<T>(IEnumerable<T> data, bool propagateExtraColumns = false);

    /// <summary>
    /// Resizes the table to the specified range address.
    /// </summary>
    /// <param name="range">The new table range.</param>
    IXlsxTable Resize(IXlsxRange range);

    /// <summary>
    /// Resizes the table to the specified range address.
    /// </summary>
    /// <param name="rangeAddress">The range boundaries.</param>
    IXlsxTable Resize(IXlsxRangeAddress rangeAddress);

    /// <summary>
    /// Resizes the table to the specified range address.
    /// </summary>
    /// <param name="rangeAddress">The range boundaries.</param>
    IXlsxTable Resize(string rangeAddress);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCell">The first cell in the range.</param>
    /// <param name="lastCell">The last cell in the range.</param>
    IXlsxTable Resize(IXlsxCell firstCell, IXlsxCell lastCell);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress">The last cell address in the worksheet.</param>
    IXlsxTable Resize(string firstCellAddress, string lastCellAddress);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress">The last cell address in the worksheet.</param>
    IXlsxTable Resize(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCellRow">The first cell's row of the range to return.</param>
    /// <param name="firstCellColumn">The first cell's column of the range to return.</param>
    /// <param name="lastCellRow">The last cell's row of the range to return.</param>
    /// <param name="lastCellColumn">The last cell's column of the range to return.</param>
    IXlsxTable Resize(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

    new IXlsxAutoFilter SetAutoFilter();

    IXlsxTable SetEmphasizeFirstColumn();

    IXlsxTable SetEmphasizeFirstColumn(bool value);

    IXlsxTable SetEmphasizeLastColumn();

    IXlsxTable SetEmphasizeLastColumn(bool value);

    IXlsxTable SetShowAutoFilter();

    IXlsxTable SetShowAutoFilter(bool value);

    IXlsxTable SetShowColumnStripes();

    IXlsxTable SetShowColumnStripes(bool value);

    IXlsxTable SetShowHeaderRow();

    IXlsxTable SetShowHeaderRow(bool value);

    IXlsxTable SetShowRowStripes();

    IXlsxTable SetShowRowStripes(bool value);

    IXlsxTable SetShowTotalsRow();

    IXlsxTable SetShowTotalsRow(bool value);

    IXlsxRangeRow TotalsRow();

    /// <summary>
    /// Converts the table to an enumerable of dynamic objects
    /// </summary>
    IEnumerable<dynamic> AsDynamicEnumerable();

    /// <summary>
    /// Converts the table to a standard .NET System.Data.DataTable
    /// </summary>
    DataTable AsNativeDataTable();

    IXlsxTable CopyTo(IXlsxWorksheet targetSheet);
}
