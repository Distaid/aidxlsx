#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Extensions;

namespace Aid.Xlsx.InsertData
{
    internal class InsertDataReaderFactory
    {
        private static readonly Lazy<InsertDataReaderFactory> _instance =
            new Lazy<InsertDataReaderFactory>(() => new InsertDataReaderFactory());

        public static InsertDataReaderFactory Instance => _instance.Value;

        [RequiresUnreferencedCode("Inserting data from an arbitrary IEnumerable uses reflection to discover the item type and its members. Use the DataTable overloads or explicit rows to stay trim-compatible.")]
        public IInsertDataReader CreateReader(IEnumerable data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            return CreateReader(data, data.GetItemType());
        }

        [RequiresUnreferencedCode("Inserting data from an arbitrary IEnumerable uses reflection to discover the item type and its members. Use the DataTable overloads or explicit rows to stay trim-compatible.")]
        internal IInsertDataReader CreateReader(IEnumerable data, Type itemType)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            if (itemType == null || itemType == typeof(Object))
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
            if (data == null) throw new ArgumentNullException(nameof(data));

            return new ArrayReader(data);
        }

        [RequiresUnreferencedCode("Inserting data from an IEnumerable of rows uses reflection to discover the row element type. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
        public IInsertDataReader CreateReader(IEnumerable<IEnumerable> data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            if (data.GetType().GetElementType() == typeof(String))
                return new SimpleTypeReader(data);

            return new ArrayReader(data);
        }

        public IInsertDataReader CreateReader(DataTable dataTable)
        {
            if (dataTable == null) throw new ArgumentNullException(nameof(dataTable));

            return new DataTableReader(dataTable);
        }
    }
}
