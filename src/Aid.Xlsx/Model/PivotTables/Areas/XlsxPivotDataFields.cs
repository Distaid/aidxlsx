using System.Collections;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.PivotTables.PivotValues;

namespace Aid.Xlsx.Model.PivotTables.Areas;

/// <summary>
/// A collection of <see cref="XlsxPivotDataField"/>.
/// </summary>
internal class XlsxPivotDataFields : IXlsxPivotValues, IReadOnlyCollection<XlsxPivotDataField>
{
    private readonly XlsxPivotTable _pivotTable;

    /// <summary>
    /// Fields displayed in the data area of the pivot table, in the order fields are displayed.
    /// </summary>
    private readonly List<XlsxPivotDataField> _fields = [];

    internal XlsxPivotDataFields(XlsxPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    public int Count => _fields.Count;

    #region IXlsxPivotValues

    public IXlsxPivotValue Add(string sourceName)
    {
        return AddField(sourceName, sourceName);
    }

    public IXlsxPivotValue Add(string sourceName, string customName)
    {
        return AddField(sourceName, customName);
    }

    public void Clear()
    {
        _fields.Clear();
        foreach (XlsxPivotDataField field in _fields)
            _pivotTable.RemoveFieldFromAxis(field.Field);
    }

    public bool Contains(string customName)
    {
        return IndexOf(customName) != -1;
    }

    public bool Contains(IXlsxPivotValue pivotValue)
    {
        return Contains(pivotValue.CustomName);
    }

    public IXlsxPivotValue Get(string customName)
    {
        XlsxPivotDataField dataField = _fields.SingleOrDefault(x => XlsxHelper.NameComparer.Equals(x.CustomName, customName)) ?? throw new KeyNotFoundException($"Unable to find data field for '{customName}'.");
        return dataField;
    }

    public IXlsxPivotValue Get(int index)
    {
        return _fields[index];
    }

    public int IndexOf(string customName)
    {
        return _fields.FindIndex(x => XlsxHelper.NameComparer.Equals(x.CustomName, customName));
    }

    public int IndexOf(IXlsxPivotValue pivotValue)
    {
        return IndexOf(pivotValue.CustomName);
    }

    public void Remove(string customName)
    {
        var index = IndexOf(customName);
        if (index == -1)
            return;

        XlsxPivotDataField dataField = _fields[index];
        _pivotTable.RemoveFieldFromAxis(dataField.Field);
        _fields.Remove(dataField);
    }

    IEnumerator<IXlsxPivotValue> IEnumerable<IXlsxPivotValue>.GetEnumerator()
    {
        return GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    internal XlsxPivotDataField AddField(string sourceName, string? customName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out FieldIndex fieldIndex))
        {
            var validNames = string.Join("','", _pivotTable.PivotCache.FieldNames);
            throw new ArgumentOutOfRangeException(nameof(sourceName), $"Field '{sourceName}' is not in the fields of a pivot cache. Should be one of '{validNames}'.");
        }

        if (fieldIndex.IsDataField)
            throw new ArgumentException("'Values' field can be used only on row or column axis.");

        var dataField = new XlsxPivotDataField(_pivotTable, fieldIndex.Value)
        {
            DataFieldName = customName,
        };
        AddField(dataField);

        // If there are multiple values, at least axis must contain 'data' field.
        // Otherwise, Excel requires a repair.
        if (_fields.Count > 1 &&
            !_pivotTable.RowAxis.ContainsDataField &&
            !_pivotTable.ColumnAxis.ContainsDataField)
        {
            _pivotTable.ColumnLabels.Add(XlsxConstants.PivotTable.ValuesSentinalLabel);
        }

        return dataField;
    }

    internal void AddField(XlsxPivotDataField dataField)
    {
        // Excel invariant - data field must have the flag if and only if it is in the data fields collection.
        _fields.Add(dataField);
        _pivotTable.PivotFields[dataField.Field].DataField = true;
    }

    public IEnumerator<XlsxPivotDataField> GetEnumerator()
    {
        return _fields.GetEnumerator();
    }
}
