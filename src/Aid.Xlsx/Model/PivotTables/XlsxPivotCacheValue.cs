using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Cells;

namespace Aid.Xlsx.PivotTables
{
    /// <summary>
    /// Represents a single value in a pivot cache record.
    /// </summary>
    internal readonly struct XlsxPivotCacheValue
    {
        /// <summary>
        /// A memory used to hold value of a <see cref="Type"/>. Its
        /// interpretation depends on the type. It doesn't hold value
        /// for strings directly, because GC doesn't allow aliasing
        /// same 8 bytes for number or references. For strings, it contains
        /// an index to a string storage array that is stored separately.
        /// </summary>
        private readonly double _value;

        private XlsxPivotCacheValue(XlsxPivotCacheValueType type, double value)
        {
            Type = type;
            _value = value;
        }

        internal XlsxPivotCacheValueType Type { get; }

        internal static XlsxPivotCacheValue ForMissing()
        {
            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.Missing, 0);
        }

        internal static XlsxPivotCacheValue ForNumber(double number)
        {
            if (double.IsNaN(number) || double.IsInfinity(number))
                throw new ArgumentOutOfRangeException();

            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.Number, number);
        }

        internal static XlsxPivotCacheValue ForBoolean(bool boolean)
        {
            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.Boolean, boolean ? 1 : 0);
        }

        internal static XlsxPivotCacheValue ForError(XlsxError error)
        {
            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.Error, (int)error);
        }

        internal static XlsxPivotCacheValue ForText(string text, List<string> storage)
        {
            var index = storage.Count;
            storage.Add(text);
            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.String, BitConverter.Int64BitsToDouble(index));
        }

        internal static XlsxPivotCacheValue ForText(string text, Dictionary<string, int> stringMap, List<string> storage)
        {
            if (!stringMap.TryGetValue(text, out var index))
            {
                index = storage.Count;
                storage.Add(text);
                stringMap.Add(text, index);
                return new XlsxPivotCacheValue(XlsxPivotCacheValueType.String, BitConverter.Int64BitsToDouble(index));
            }

            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.String, BitConverter.Int64BitsToDouble(index));
        }

        internal static XlsxPivotCacheValue ForDateTime(DateTime dateTime)
        {
            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.DateTime, BitConverter.Int64BitsToDouble(dateTime.Ticks));
        }

        internal static XlsxPivotCacheValue ForIndex(uint index)
        {
            return new XlsxPivotCacheValue(XlsxPivotCacheValueType.Index, BitConverter.Int64BitsToDouble(index));
        }

        internal XlsxCellValue GetCellValue(List<string> stringStorage, XlsxPivotCacheSharedItems sharedItems)
        {
            switch (Type)
            {
                case XlsxPivotCacheValueType.Missing:
                    return Blank.Value;

                case XlsxPivotCacheValueType.Number:
                    return _value;

                case XlsxPivotCacheValueType.Boolean:
                    return _value != 0;

                case XlsxPivotCacheValueType.Error:
                    return (XlsxError)_value;

                case XlsxPivotCacheValueType.String:
                    return GetText(stringStorage);

                case XlsxPivotCacheValueType.DateTime:
                    return GetDateTime();

                case XlsxPivotCacheValueType.Index:
                    var intIndex = unchecked((uint)BitConverter.DoubleToInt64Bits(_value));
                    return sharedItems[intIndex];

                default:
                    throw new NotSupportedException();
            }
        }

        internal double GetNumber() => _value;

        internal Boolean GetBoolean()
        {
            return _value != 0;
        }

        internal XlsxError GetError()
        {
            return (XlsxError)_value;
        }

        internal string GetText(IReadOnlyList<string> stringStorage)
        {
            var stringIndex = unchecked((int)BitConverter.DoubleToInt64Bits(_value));
            return stringStorage[stringIndex];
        }

        internal DateTime GetDateTime()
        {
            var ticks = BitConverter.DoubleToInt64Bits(_value);
            return new DateTime(ticks);
        }

        internal uint GetIndex()
        {
            return unchecked((uint)BitConverter.DoubleToInt64Bits(_value));
        }
    }
}
