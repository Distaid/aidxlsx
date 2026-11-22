using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace Aid.Xlsx.Worksheets
{
    public interface IXlsxWorksheets : IEnumerable<IXlsxWorksheet>
    {
        int Count { get; }

        IXlsxWorksheet Add();

        IXlsxWorksheet Add(Int32 position);

        IXlsxWorksheet Add(String sheetName);

        IXlsxWorksheet Add(String sheetName, Int32 position);

        IXlsxWorksheet Add(DataTable dataTable);

        IXlsxWorksheet Add(DataTable dataTable, String sheetName);

        IXlsxWorksheet Add(DataTable dataTable, String sheetName, String tableName);

        void Add(DataSet dataSet);

        Boolean Contains(String sheetName);

        void Delete(String sheetName);

        void Delete(Int32 position);

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

        IXlsxWorksheet Worksheet(Int32 position);
    }
}
