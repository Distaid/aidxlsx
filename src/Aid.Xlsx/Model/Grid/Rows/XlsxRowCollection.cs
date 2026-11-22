#nullable disable

using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Model.Grid.Rows;

using System.Collections;

internal class XlsxRowsCollection : IDictionary<Int32, XlsxRow>
{
    private readonly Dictionary<Int32, XlsxRow> _dictionary = new Dictionary<Int32, XlsxRow>();

    public Dictionary<Int32, XlsxRow> Deleted { get; } = new Dictionary<Int32, XlsxRow>();

    public Int32 MaxRowUsed;

    #region IDictionary<int,XlsxRow> Members

    public void Add(int key, XlsxRow value)
    {
        if (key > MaxRowUsed) MaxRowUsed = key;

        Deleted.Remove(key);
        _dictionary.Add(key, value);
    }

    public bool ContainsKey(int key)
    {
        return _dictionary.ContainsKey(key);
    }

    public ICollection<int> Keys
    {
        get { return _dictionary.Keys; }
    }

    public bool Remove(int key)
    {
        if (!Deleted.ContainsKey(key))
            Deleted.Add(key, _dictionary[key]);

        return _dictionary.Remove(key);
    }

    public bool TryGetValue(int key, out XlsxRow value)
    {
        return _dictionary.TryGetValue(key, out value);
    }

    public ICollection<XlsxRow> Values
    {
        get { return _dictionary.Values; }
    }

    public XlsxRow this[int key]
    {
        get { return _dictionary[key]; }
        set { _dictionary[key] = value; }
    }

    public void Add(KeyValuePair<int, XlsxRow> item)
    {
        if (item.Key > MaxRowUsed) MaxRowUsed = item.Key;

        Deleted.Remove(item.Key);
        _dictionary.Add(item.Key, item.Value);
    }

    public void Clear()
    {
        _dictionary.Clear();
    }

    public bool Contains(KeyValuePair<int, XlsxRow> item)
    {
        return _dictionary.Contains(item);
    }

    public void CopyTo(KeyValuePair<int, XlsxRow>[] array, int arrayIndex)
    {
        throw new NotImplementedException();
    }

    public int Count
    {
        get { return _dictionary.Count; }
    }

    public bool IsReadOnly
    {
        get { return false; }
    }

    public bool Remove(KeyValuePair<int, XlsxRow> item)
    {
        if (!Deleted.ContainsKey(item.Key))
            Deleted.Add(item.Key, _dictionary[item.Key]);

        return _dictionary.Remove(item.Key);
    }

    public IEnumerator<KeyValuePair<int, XlsxRow>> GetEnumerator()
    {
        return _dictionary.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return _dictionary.GetEnumerator();
    }

    #endregion IDictionary<int,XlsxRow> Members

    public void ShiftRowsDown(Int32 startingRow, Int32 rowsToShift)
    {
        foreach (int ro in _dictionary.Keys.Where(k => k >= startingRow).OrderByDescending(k => k))
        {
            var rowToMove = _dictionary[ro];
            _dictionary.Remove(ro);
            Int32 newRowNum = ro + rowsToShift;
            if (newRowNum <= XlsxHelper.MaxRowNumber)
            {
                rowToMove.SetRowNumber(newRowNum);
                _dictionary.Add(newRowNum, rowToMove);
            }
        }
    }

    public void RemoveAll(Func<XlsxRow, Boolean> predicate)
    {
        foreach (var row in _dictionary.Values.Where(predicate).Where(row1 => !Deleted.ContainsKey(row1.RowNumber())))
        {
            Deleted.Add(row.RowNumber(), row);
        }

        _dictionary.RemoveAll(predicate);
    }
}
