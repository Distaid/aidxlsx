// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.DataImport.RowReaders;

/// <summary>
/// A universal interface for different data readers used in InsertData logic.
/// </summary>
internal interface IInsertDataReader
{
    /// <summary>
    /// Get a collection of records, each as a collection of values, extracted from a source.
    /// </summary>
    IEnumerable<IEnumerable<XlsxCellValue>> GetRecords();

    /// <summary>
    /// Get the number of properties to use as a table with.
    /// Actual number of may vary in different records.
    /// </summary>
    int GetPropertiesCount();

    /// <summary>
    /// Get the title of the property with the specified index.
    /// </summary>
    string? GetPropertyName(int propertyIndex);

    /// <summary>Gets the number format for data cells, if the reader supplies one.</summary>
    string? GetNumberFormat(int propertyIndex) => null;
}
