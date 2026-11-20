#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Aid.Xlsx.InsertData
{
    internal class SimpleTypeReader : IInsertDataReader
    {
        private readonly IEnumerable<object> _data;
        private readonly Type _itemType;

        [RequiresUnreferencedCode("The item type of the sequence is discovered with reflection. Pass a typed sequence or a DataTable to stay trim-compatible.")]
        public SimpleTypeReader(IEnumerable data)
            : this(data, data.GetItemType())
        {
        }

        /// <summary>
        /// Create a reader for simple values of a known item type. The type is passed explicitly,
        /// because <paramref name="data"/> can be a non-generic projection (e.g. <c>List&lt;object&gt;</c>).
        /// </summary>
        internal SimpleTypeReader(IEnumerable data, Type itemType)
        {
            _data = data?.Cast<object>() ?? throw new ArgumentNullException(nameof(data));
            _itemType = itemType;
        }

        public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
        {
            return _data.Select(item => new[] { item }.Select(XLCellValue.FromInsertedObject));
        }

        public int GetPropertiesCount()
        {
            return 1;
        }

        public string GetPropertyName(int propertyIndex = 0)
        {
            if (propertyIndex != 0)
                throw new ArgumentException("SimpleTypeReader supports only a single property");

            return _itemType.Name;
        }
    }
}
