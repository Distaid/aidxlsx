#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace Aid.Xlsx.InsertData
{
    internal class SimpleNullableTypeReader : IInsertDataReader
    {
        private readonly IEnumerable<object> _data;
        private readonly Type _itemType;

        [RequiresUnreferencedCode("The item type of the nullable sequence is discovered with reflection. Pass a typed sequence or a DataTable to stay trim-compatible.")]
        public SimpleNullableTypeReader(IEnumerable data)
        {
            _data = data?.Cast<object>() ?? throw new ArgumentNullException(nameof(data));
            _itemType = data.GetItemType().GetUnderlyingType();
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
                throw new ArgumentException("SimpleNullableTypeReader supports only a single property");

            return _itemType.Name;
        }
    }
}
