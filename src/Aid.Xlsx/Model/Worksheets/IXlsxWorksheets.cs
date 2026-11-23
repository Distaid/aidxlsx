using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace Aid.Xlsx.Model.Worksheets;

public interface IXlsxWorksheets : IEnumerable<IXlsxWorksheet>
{
    int Count { get; }

    IXlsxWorksheet Add();

    IXlsxWorksheet Add(int position);

    IXlsxWorksheet Add(string sheetName);

    IXlsxWorksheet Add(string sheetName, int position);

    IXlsxWorksheet Add(DataTable dataTable);

    IXlsxWorksheet Add(DataTable dataTable, string sheetName);

    IXlsxWorksheet Add(DataTable dataTable, string sheetName, string tableName);

    void Add(DataSet dataSet);

    bool Contains(string sheetName);

    void Delete(string sheetName);

    void Delete(int position);

    /// <summary>
    /// Try to get a sheet of a workbook with the specified name. Sheet names are case-insensitive.
    /// </summary>
    /// <param name="sheetName">Name of sought sheet.</param>
    /// <param name="worksheet">Found sheet or null if sheet is not found.</param>
    /// <returns><c>true</c> when sheet was found or <c>false</c> when it wasn't.</returns>
    bool TryGetWorksheet(string sheetName, [NotNullWhen(true)] out IXlsxWorksheet? worksheet);

    /// <summary>
    /// Get a sheet of a workbook with specified name. Sheet names are case-insensitive.
    /// </summary>
    /// <param name="sheetName">Name of sought sheet.</param>
    /// <returns>Sheet with the specified name.</returns>
    /// <exception cref="KeyNotFoundException">When sheet with <paramref name="sheetName"/> isn't among the sheets.</exception>
    IXlsxWorksheet Worksheet(string sheetName);

    IXlsxWorksheet Worksheet(int position);
}
