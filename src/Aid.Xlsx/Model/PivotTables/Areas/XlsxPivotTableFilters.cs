// Keep this file CodeMaid organised and cleaned
using System.Collections;
using System.Diagnostics;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Model.PivotTables.Areas;

/// <summary>
/// Page/filter fields of a <see cref="XlsxPivotTable"/>. It determines filter values and layout.
/// It is accessible through fluent API <see cref="XlsxPivotTable.ReportFilters"/>.
/// </summary>
internal class XlsxPivotTableFilters : IXlsxPivotFields
{
    private readonly XlsxPivotTable _pivotTable;

    /// <summary>
    /// Filter fields in correct order. The layout is determined by
    /// <see cref="XlsxPivotTable.FilterFieldsPageWrap"/> and
    /// <see cref="XlsxPivotTable.FilterAreaOrder"/>.
    /// </summary>
    private readonly List<XlsxPivotPageField> _fields = new();

    internal XlsxPivotTableFilters(XlsxPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    IXlsxPivotField IXlsxPivotFields.Add(String sourceName) => Add(sourceName, sourceName);

    IXlsxPivotField IXlsxPivotFields.Add(String sourceName, String customName) => Add(sourceName, customName);

    public void Clear()
    {
        foreach (var field in _fields)
            _pivotTable.RemoveFieldFromAxis(field.Field);

        _fields.Clear();
    }

    public Boolean Contains(String sourceName)
    {
        return IndexOf(sourceName) >= 0;
    }

    public bool Contains(IXlsxPivotField pivotField)
    {
        return Contains(pivotField.SourceName);
    }

    public IXlsxPivotField Get(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XlsxPivotAxis.AxisPage}.");

        var filterField = _fields.SingleOrDefault(f => f.Field == fieldIndex);
        if (filterField is null)
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XlsxPivotAxis.AxisPage}.");

        return new XlsxPivotTablePageField(_pivotTable, filterField);
    }

    public IXlsxPivotField Get(Int32 index)
    {
        if (index < 0 || index >= _fields.Count)
            throw new IndexOutOfRangeException();

        return new XlsxPivotTablePageField(_pivotTable, _fields[index]);
    }

    IEnumerator<IXlsxPivotField> IEnumerable<IXlsxPivotField>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerator<XlsxPivotTablePageField> GetEnumerator()
    {
        foreach (var field in _fields)
            yield return new XlsxPivotTablePageField(_pivotTable, field);
    }

    public Int32 IndexOf(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            return -1;

        return _fields.FindIndex(f => f.Field == fieldIndex);
    }

    public Int32 IndexOf(IXlsxPivotField pf)
    {
        return IndexOf(pf.SourceName);
    }

    public void Remove(String sourceName)
    {
        var index = IndexOf(sourceName);
        if (index == -1)
            return;

        var heightDifference = GetHeightDifference(-1);
        var movedArea = _pivotTable.Area.ShiftRows(heightDifference);

        _fields.RemoveAt(index);
        _pivotTable.RemoveFieldFromAxis(index);

        _pivotTable.Area = movedArea;
    }

    internal IReadOnlyList<XlsxPivotPageField> Fields => _fields;

    internal XlsxPivotTablePageField Add(String sourceName, String customName)
    {
        if (sourceName == XlsxConstants.PivotTable.ValuesSentinalLabel)
            throw new ArgumentException(nameof(sourceName), $"The column '{sourceName}' does not appear in the source range.");

        var heightDifference = GetHeightDifference(1);
        var movedArea = _pivotTable.Area.ShiftRows(heightDifference);

        var fieldIndex = _pivotTable.AddFieldToAxis(sourceName, customName, XlsxPivotAxis.AxisPage);
        var filterField = new XlsxPivotPageField(fieldIndex);
        _fields.Add(filterField);

        _pivotTable.Area = movedArea;
        return new XlsxPivotTablePageField(_pivotTable, filterField);
    }

    internal bool Contains(FieldIndex fieldIndex)
    {
        return _fields.FindIndex(f => f.Field == fieldIndex) >= 0;
    }

    internal void AddField(XlsxPivotPageField pageField)
    {
        _fields.Add(pageField);
    }

    /// <summary>
    /// Number of rows/cols occupied by the filter area. Filter area is above the pivot table and it
    /// optional (i.e. size <c>0</c> indicates no filter).
    /// </summary>
    internal (int Width, int Height) GetSize()
    {
        return GetSize(_fields.Count, _pivotTable.FilterAreaOrder, _pivotTable.FilterFieldsPageWrap);
    }

    /// <summary>
    /// Number of rows/cols occupied by the filter area, including the gap below, if there is at least one filter.
    /// </summary>
    internal (int Width, int Height) GetSizeWithGap()
    {
        return GetSizeWithGap(_fields.Count, _pivotTable.FilterAreaOrder, _pivotTable.FilterFieldsPageWrap);
    }

    private int GetHeightDifference(int fieldChangeCount)
    {
        var originalHeight = GetSizeWithGap(_fields.Count, _pivotTable.FilterAreaOrder, _pivotTable.FilterFieldsPageWrap).Height;
        var modifiedHeight = GetSizeWithGap(_fields.Count + fieldChangeCount, _pivotTable.FilterAreaOrder, _pivotTable.FilterFieldsPageWrap).Height;
        return modifiedHeight - originalHeight;
    }

    private static (int Width, int Height) GetSize(int fieldCount, XlsxFilterAreaOrder order, int filterWrap)
    {
        if (filterWrap == 0)
            filterWrap = int.MaxValue;

        var dim1 = Math.DivRem(fieldCount, filterWrap, out var dim2);
        dim1 = fieldCount > 0 ? dim1 + 1 : dim1;

        return order switch
        {
            XlsxFilterAreaOrder.DownThenOver => new(dim1, dim2),
            XlsxFilterAreaOrder.OverThenDown => new(dim2, dim1),
            _ => throw new UnreachableException(),
        };
    }

    private static (int Width, int Height) GetSizeWithGap(int fieldCount, XlsxFilterAreaOrder order, int filterWrap)
    {
        var filtersSize = GetSize(fieldCount, order, filterWrap);
        return filtersSize.Height > 0
            ? (filtersSize.Width, filtersSize.Height + 1)
            : filtersSize;
    }
}
