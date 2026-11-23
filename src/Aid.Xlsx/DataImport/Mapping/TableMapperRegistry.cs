using System.ComponentModel;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.DataImport.RowReaders;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.DataImport.Mapping;

/// <summary>Registers row mappers emitted by the Aid.Xlsx source generator.</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class TableMapperRegistry
{
    /// <summary>Registers a mapper for <typeparamref name="T"/>.</summary>
    public static void Register<T>(string? tableName, string[] headers, string?[] formats, Func<T, int, XlsxCellValue> read)
    {
        ArgumentNullException.ThrowIfNull(headers);
        ArgumentNullException.ThrowIfNull(formats);
        ArgumentNullException.ThrowIfNull(read);
        if (headers.Length == 0 || headers.Length != formats.Length)
            throw new ArgumentException("Provide a nonempty header list and one format per header.", nameof(headers));

        var mapper = new TableRowMap<T>(tableName, (string[])headers.Clone(), (string?[])formats.Clone(), read);
        if (Interlocked.CompareExchange(ref TableRowMap<T>.Registered, mapper, null) is not null)
            throw new InvalidOperationException($"A table mapper for '{typeof(T)}' is already registered.");
    }
}

internal sealed record TableRowMap<T>(string? TableName, string[] Headers, string?[] Formats, Func<T, int, XlsxCellValue> Read)
{
    internal static TableRowMap<T>? Registered;

    internal static TableRowMap<T> GetRegistered() => Volatile.Read(ref Registered)
        ?? throw new NotSupportedException($"Add [XlsxTable] to '{typeof(T)}' and enable the Aid.Xlsx source generator.");
}

internal sealed class MappedDataReader<T>(IEnumerable<T> data, TableRowMap<T> mapper) : IInsertDataReader
{
    public int GetPropertiesCount() => mapper.Headers.Length;

    public string GetPropertyName(int propertyIndex) => mapper.Headers[propertyIndex];

    public string? GetNumberFormat(int propertyIndex) => mapper.Formats[propertyIndex];

    public IEnumerable<IEnumerable<XlsxCellValue>> GetRecords()
    {
        foreach (T? item in data)
            yield return ReadRow(item);
    }

    private IEnumerable<XlsxCellValue> ReadRow(T item)
    {
        for (var i = 0; i < mapper.Headers.Length; i++)
            yield return item is null ? Blank.Value : mapper.Read(item, i);
    }
}
