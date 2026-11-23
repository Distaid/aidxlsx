using System.Diagnostics;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// <para>
/// A list of <see cref="XlsxPivotCacheValue"/> in the pivot table cache
/// definition. Generally, it contains all strings of the field records
/// (record just indexes them through <see cref="XlsxPivotCacheValueType.Index"/>)
/// and also values used directly in pivot table (e.g. filter field reference
/// the table definition, not record).
/// </para>
/// <para>
/// Shared items can't contain <see cref="XlsxPivotCacheValueType.Index"/>.
/// </para>
/// </summary>
internal class XlsxPivotCacheSharedItems
{
    private readonly List<XlsxPivotCacheValue> _values = [];

    /// <summary>
    /// Storage of strings to save 8 bytes per <c>XlsxPivotCacheValue</c>
    /// (reference can't be aliased with a number).
    /// </summary>
    private readonly List<string> _stringStorage = [];

    /// <summary>
    /// Strings in a pivot table are case-insensitive.
    /// </summary>
    private readonly Dictionary<string, int> _stringMap = [with(StringComparer.OrdinalIgnoreCase)];

    internal XlsxCellValue this[uint index] => GetValue(index).GetCellValue(_stringStorage, this);

    internal int Count => _values.Count;

    internal void Add(XlsxCellValue value)
    {
        switch (value.Type)
        {
            case XlsxDataType.Blank:
                AddMissing();
                break;
            case XlsxDataType.Boolean:
                AddBoolean(value.GetBoolean());
                break;
            case XlsxDataType.Number:
                AddNumber(value.GetNumber());
                break;
            case XlsxDataType.Text:
                AddString(value.GetText());
                break;
            case XlsxDataType.Error:
                AddError(value.GetError());
                break;
            case XlsxDataType.DateTime:
                AddDateTime(value.GetDateTime());
                break;
            case XlsxDataType.TimeSpan:
                DateTime timeSpan = value.GetTimeSpan().ToSerialDateTime().ToSerialDateTime();
                AddDateTime(timeSpan);
                break;
            default:
                throw new UnreachableException();
        }
    }

    internal void AddMissing()
    {
        _values.Add(XlsxPivotCacheValue.ForMissing());
    }

    internal void AddNumber(double number)
    {
        _values.Add(XlsxPivotCacheValue.ForNumber(number));
    }

    internal void AddBoolean(bool boolean)
    {
        _values.Add(XlsxPivotCacheValue.ForBoolean(boolean));
    }

    internal void AddError(XlsxError error)
    {
        _values.Add(XlsxPivotCacheValue.ForError(error));
    }

    internal void AddString(string text)
    {
        // Shared items doesn't distinguish between two texts that differ only in case.
        if (!_stringMap.ContainsKey(text))
        {
            var index = _stringStorage.Count;
            _values.Add(XlsxPivotCacheValue.ForText(text, _stringStorage));
            _stringMap.Add(text, index);
        }
    }

    internal void AddDateTime(DateTime dateTime)
    {
        _values.Add(XlsxPivotCacheValue.ForDateTime(dateTime));
    }

    internal IEnumerable<XlsxCellValue> GetCellValues()
    {
        foreach (XlsxPivotCacheValue value in _values)
        {
            yield return value.GetCellValue(_stringStorage, this);
        }
    }

    internal XlsxPivotCacheValue GetValue(uint index)
    {
        return _values[checked((int)index)];
    }

    internal string GetStringValue(uint index)
    {
        XlsxPivotCacheValue value = GetValue(index);
        return value.GetText(_stringStorage);
    }

    /// <summary>
    /// Get index of value or -1 if not among shared items.
    /// </summary>
    internal int IndexOf(XlsxCellValue value)
    {
        for (var index = 0; index < _values.Count; ++index)
        {
            XlsxPivotCacheValue sharedValue = _values[index];
            XlsxCellValue cacheValue = sharedValue.GetCellValue(_stringStorage, this);
            if (XlsxCellValueComparer.OrdinalIgnoreCase.Equals(cacheValue, value))
                return index;
        }

        return -1;
    }
}
