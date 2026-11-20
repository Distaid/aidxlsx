#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace Aid.Xlsx.InsertData
{
    internal class UntypedObjectReader : IInsertDataReader
    {
        private readonly IEnumerable<object> _data;
        private readonly IEnumerable<IInsertDataReader> _readers;

        [RequiresUnreferencedCode("A sequence with mixed runtime item types is split by type and dispatched to reflection-based readers. Use a homogeneous typed sequence or a DataTable to stay trim-compatible.")]
        public UntypedObjectReader(IEnumerable data)
        {
            _data = (data ?? new object[0]).Cast<object>();
            _readers = CreateReaders().ToList();

            IEnumerable<IInsertDataReader> CreateReaders()
            {
                if (!_data.Any())
                    yield break;

                List<object> itemsOfSameType = new List<object>();
                Type previousType = null;

                foreach (var item in _data)
                {
                    var currentType = item?.GetType();

                    if (previousType != currentType && itemsOfSameType.Count > 0)
                    {
                        yield return CreateReader(itemsOfSameType, previousType);
                        itemsOfSameType.Clear();
                    }
                    itemsOfSameType.Add(item);
                    previousType = currentType;
                }

                if (itemsOfSameType.Count > 0)
                {
                    yield return CreateReader(itemsOfSameType, previousType);
                }
            }

            IInsertDataReader CreateReader(List<object> itemsOfSameType, Type itemType)
            {
                // Snapshot the group: the source list is reused for the next group of items, while
                // the created reader enumerates its data lazily. Without a copy the reader would see
                // the items of the following groups (with possibly different runtime types).
                var items = new List<object>(itemsOfSameType);

                if (itemType == null)
                    return new NullDataReader(items);

                return InsertDataReaderFactory.Instance.CreateReader(items, itemType);
            }
        }

        public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
        {
            foreach (var reader in _readers)
            {
                foreach (var item in reader.GetRecords())
                {
                    yield return item;
                }
            }
        }

        public int GetPropertiesCount()
        {
            return GetFirstNonNullReader()?.GetPropertiesCount() ?? 0;
        }

        public string GetPropertyName(int propertyIndex)
        {
            return GetFirstNonNullReader()?.GetPropertyName(propertyIndex);
        }

        private IInsertDataReader GetFirstNonNullReader()
        {
            return _readers.FirstOrDefault(r => !(r is NullDataReader));
        }
    }
}
