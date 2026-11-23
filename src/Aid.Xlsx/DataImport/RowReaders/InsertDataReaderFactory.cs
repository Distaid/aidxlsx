#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.DataImport.RowReaders;

internal class InsertDataReaderFactory
{
    private static readonly Lazy<InsertDataReaderFactory> _instance =
        new(() => new InsertDataReaderFactory());

    public static InsertDataReaderFactory Instance => _instance.Value;

    [RequiresUnreferencedCode("Inserting data from an arbitrary IEnumerable uses reflection to discover the item type and its members. Use the DataTable overloads or explicit rows to stay trim-compatible.")]
    public IInsertDataReader CreateReader(IEnumerable data)
    {
        ArgumentNullException.ThrowIfNull(data);

        return CreateReader(data, data.GetItemType());
    }

    [RequiresUnreferencedCode("Inserting data from an arbitrary IEnumerable uses reflection to discover the item type and its members. Use the DataTable overloads or explicit rows to stay trim-compatible.")]
    internal IInsertDataReader CreateReader(IEnumerable data, Type itemType)
    {
        ArgumentNullException.ThrowIfNull(data);

        if (itemType == null || itemType == typeof(object))
            return new UntypedObjectReader(data);
        else if (itemType.IsNullableType() && itemType.GetUnderlyingType().IsSimpleType())
            return new SimpleNullableTypeReader(data, itemType);
        else if (itemType.IsSimpleType())
            return new SimpleTypeReader(data, itemType);
        else if (typeof(IDataRecord).IsAssignableFrom(itemType))
            return new DataRecordReader(data.OfType<IDataRecord>());
        else if (itemType.IsArray || typeof(IEnumerable).IsAssignableFrom(itemType))
            return new ArrayReader(data.Cast<IEnumerable>());
        else if (itemType == typeof(DataRow))
            return new DataTableReader(data.Cast<DataRow>());

        return new ObjectReader(data, itemType);
    }

    public IInsertDataReader CreateReader<T>(IEnumerable<T[]> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        return new ArrayReader(data);
    }

    [RequiresUnreferencedCode("Inserting data from an IEnumerable of rows uses reflection to discover the row element type. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    public IInsertDataReader CreateReader(IEnumerable<IEnumerable> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        if (data.GetType().GetElementType() == typeof(string))
            return new SimpleTypeReader(data);

        return new ArrayReader(data);
    }

    public IInsertDataReader CreateReader(DataTable dataTable)
    {
        ArgumentNullException.ThrowIfNull(dataTable);

        return new DataTableReader(dataTable);
    }
}
