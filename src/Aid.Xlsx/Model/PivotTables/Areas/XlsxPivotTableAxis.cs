using System.Collections;

namespace Aid.Xlsx.PivotTables.Areas;

/// <summary>
/// A description of one axis (<see cref="XlsxPivotTable.RowAxis"/>/<see cref="XlsxPivotTable.ColumnAxis"/>)
/// of a <see cref="XlsxPivotTable"/>. It consists of fields in a specific order and values that make up
/// individual rows/columns of the axis.
/// </summary>
/// <remarks>
/// [ISO-29500] 18.10.1.17 colItems (Column Items), 18.10.1.84 rowItems (Row Items).
/// </remarks>
internal class XlsxPivotTableAxis : IXlsxPivotFields
{
    private readonly XlsxPivotTable _pivotTable;

    private readonly XlsxPivotAxis _axis;

    /// <summary>
    /// Fields displayed on the axis, in the order of the fields on the axis.
    /// </summary>
    private readonly List<FieldIndex> _fields = new();

    /// <summary>
    /// Values of one row/column in an axis. Items are not kept in sync with <see cref="_fields"/>.
    /// </summary>
    private readonly List<XlsxPivotFieldAxisItem> _axisItems = new();

    internal XlsxPivotTableAxis(XlsxPivotTable pivotTable, XlsxPivotAxis axis)
    {
        _pivotTable = pivotTable;
        _axis = axis;
    }

    /// <summary>
    /// A list of fields to displayed on the axis. It determines which fields and in what order
    /// should the fields be displayed.
    /// </summary>
    internal IReadOnlyList<FieldIndex> Fields => _fields;

    /// <summary>
    /// Individual row/column parts of the axis.
    /// </summary>
    internal IReadOnlyList<XlsxPivotFieldAxisItem> Items => _axisItems;

    internal bool ContainsDataField => _fields.Any(x => x.IsDataField);

    IXlsxPivotField IXlsxPivotFields.Add(String sourceName) => Add(sourceName, sourceName);

    IXlsxPivotField IXlsxPivotFields.Add(String sourceName, String customName) => Add(sourceName, customName);

    void IXlsxPivotFields.Clear() => Clear();

    Boolean IXlsxPivotFields.Contains(String sourceName) => Contains(sourceName);

    Boolean IXlsxPivotFields.Contains(IXlsxPivotField pivotField) => Contains(pivotField.SourceName);

    IXlsxPivotField IXlsxPivotFields.Get(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index) ||
            !_fields.Contains(index))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {_axis}.");

        return new XlsxPivotTableAxisField(_pivotTable, index);
    }

    IXlsxPivotField IXlsxPivotFields.Get(Int32 index)
    {
        if (index < 0 || index >= _fields.Count)
            throw new IndexOutOfRangeException();

        return new XlsxPivotTableAxisField(_pivotTable, _fields[index]);
    }

    Int32 IXlsxPivotFields.IndexOf(String sourceName)
    {
        return IndexOf(sourceName);
    }

    Int32 IXlsxPivotFields.IndexOf(IXlsxPivotField pf)
    {
        return IndexOf(pf.SourceName);
    }

    void IXlsxPivotFields.Remove(String sourceName)
    {
        var index = IndexOf(sourceName);
        if (index == -1)
            return;

        _pivotTable.RemoveFieldFromAxis(_fields[index]);
        _fields.RemoveAt(index);
    }

    IEnumerator<IXlsxPivotField> IEnumerable<IXlsxPivotField>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerator<XlsxPivotTableAxisField> GetEnumerator()
    {
        foreach (var fieldIndex in _fields)
            yield return new XlsxPivotTableAxisField(_pivotTable, fieldIndex);
    }

    internal int IndexOf(FieldIndex index)
    {
        return _fields.IndexOf(index);
    }

    internal bool Contains(string sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index))
            return false;

        return _fields.Contains(index);
    }

    /// <summary>
    /// Add field to the axis, as an index.
    /// </summary>
    internal void AddField(FieldIndex fieldIndex)
    {
        if (_pivotTable.IsFieldUsedOnAxis(fieldIndex))
            throw new ArgumentException("Field is already used on an axis.");

        _fields.Add(fieldIndex);
    }

    private XlsxPivotTableAxisField Add(String sourceName, String customName)
    {
        var field = AddField(sourceName, customName);

        // Excel by default adds a subtotal, but earlier versions of the port didn't have them,
        // so keep API behavior.
        if (field.Offset != FieldIndex.DataField.Value)
            _pivotTable.PivotFields[field.Offset].RemoveSubtotal(XlsxSubtotalFunction.Automatic);

        return field;
    }

    internal XlsxPivotTableAxisField AddField(String sourceName, String customName)
    {
        var index = _pivotTable.AddFieldToAxis(sourceName, customName, _axis);
        _fields.Add(index);
        return new XlsxPivotTableAxisField(_pivotTable, index);
    }

    /// <summary>
    /// Add a row/column axis values (i.e. values visible on the axis).
    /// </summary>
    internal void AddItem(XlsxPivotFieldAxisItem axisItem)
    {
        _axisItems.Add(axisItem);
    }

    internal void Clear()
    {
        foreach (var fieldIndex in _fields)
            _pivotTable.RemoveFieldFromAxis(fieldIndex);

        _axisItems.Clear();
        _fields.Clear();
    }

    private Int32 IndexOf(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            return -1;

        return _fields.IndexOf(fieldIndex);
    }
}
