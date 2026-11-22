using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.DataImport.Mapping;

/// <summary>Inserts rows using mappers generated at compile time.</summary>
public static class MappedDataExtensions
{
    /// <summary>Inserts rows mapped by <see cref="Attributes.XlsxTableAttribute"/> starting at the specified cell.</summary>
    public static IXlsxRange InsertMappedData<T>(this IXlsxCell cell, IEnumerable<T> data, bool addHeadings = false)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentNullException.ThrowIfNull(data);
        var mapper = TableRowMap<T>.GetRegistered();
        return GetWorksheet(cell).InsertData(Point.FromAddress(cell.Address), new MappedDataReader<T>(data, mapper), addHeadings, transpose: false);
    }

    /// <summary>Creates an Excel table from rows mapped by <see cref="Attributes.XlsxTableAttribute"/>.</summary>
    public static IXlsxTable InsertMappedTable<T>(this IXlsxCell cell, IEnumerable<T> data, string? tableName = null)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentNullException.ThrowIfNull(data);
        var mapper = TableRowMap<T>.GetRegistered();
        return GetWorksheet(cell).InsertTable(Point.FromAddress(cell.Address), new MappedDataReader<T>(data, mapper),
            tableName ?? mapper.TableName, createTable: true, addHeadings: true, transpose: false);
    }

    private static XlsxWorksheet GetWorksheet(IXlsxCell cell) => cell.Worksheet as XlsxWorksheet
        ?? throw new ArgumentException("The cell must belong to an Aid.Xlsx worksheet.", nameof(cell));
}
