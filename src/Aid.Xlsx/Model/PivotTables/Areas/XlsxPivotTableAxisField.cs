#nullable disable
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

namespace Aid.Xlsx.Model.PivotTables.Areas;

/// <summary>
/// A fluent API for one field in <see cref="XlsxPivotTableAxis"/>, either
/// <see cref="XlsxPivotTable.RowLabels"/> or <see cref="XlsxPivotTable.ColumnLabels"/>.
/// </summary>
internal class XlsxPivotTableAxisField : IXlsxPivotField
{
    private readonly XlsxPivotTable _pivotTable;
    private readonly FieldIndex _index;

    internal XlsxPivotTableAxisField(XlsxPivotTable pivotTable, FieldIndex index)
    {
        _pivotTable = pivotTable;
        _index = index;
    }

    #region IXlsxPivotField memebers

    public string SourceName
    {
        get
        {
            if (_index.IsDataField)
                return XlsxConstants.PivotTable.ValuesSentinalLabel;

            return _pivotTable.PivotCache.FieldNames[_index];
        }
    }

    public string CustomName
    {
        get => GetFieldValue(f => f.Name, _pivotTable.DataCaption);
        set
        {
            if (_index.IsDataField)
            {
                _pivotTable.DataCaption = value;
                return;
            }

            if (_pivotTable.TryGetCustomNameFieldIndex(value, out FieldIndex idx) && idx != _index)
                throw new ArgumentException($"Custom name '{value}' is already used by another field.");

            _pivotTable.PivotFields[_index].Name = value;
        }
    }

    public string SubtotalCaption
    {
        get => GetFieldValue(f => f.SubtotalCaption, string.Empty);
        set => GetField().SubtotalCaption = value;
    }

    public IReadOnlyCollection<XlsxSubtotalFunction> Subtotals
    {
        get
        {
            HashSet<XlsxSubtotalFunction> subtotal = GetField().Subtotals;
            var isCustomSubtotal = subtotal.Count > 1 || (subtotal.Count > 0 && !subtotal.Contains(XlsxSubtotalFunction.Automatic));
            if (isCustomSubtotal)
            {
                // When subtotal is custom, the automatic is not shown
                subtotal = [.. subtotal];
                subtotal.Remove(XlsxSubtotalFunction.Automatic);
            }

            return subtotal;
        }
    }

    public bool IncludeNewItemsInFilter
    {
        get => GetFieldValue(f => f.IncludeNewItemsInFilter, false);
        set => GetField().IncludeNewItemsInFilter = value;
    }

    public bool Outline
    {
        get => GetFieldValue(f => f.Outline, true);
        set => GetField().Outline = value;
    }
    public bool Compact
    {
        get => GetFieldValue(f => f.Compact, true);
        set => GetField().Compact = value;
    }

    public bool? SubtotalsAtTop
    {
        get => GetFieldValue(f => f.SubtotalTop, true);
        set => GetField().SubtotalTop = value ?? true;
    }

    public bool RepeatItemLabels
    {
        get => GetFieldValue(f => f.RepeatItemLabels, false);
        set => GetField().RepeatItemLabels = value;
    }

    public bool InsertBlankLines
    {
        get => GetFieldValue(f => f.InsertBlankRow, false);
        set => GetField().InsertBlankRow = value;
    }

    public bool ShowBlankItems
    {
        get => GetFieldValue(f => f.ShowAll, true);
        set => GetField().ShowAll = value;
    }

    public bool InsertPageBreaks
    {
        get => GetFieldValue(f => f.InsertPageBreak, false);
        set => GetField().InsertPageBreak = value;
    }

    public bool Collapsed
    {
        get
        {
            return GetFieldValue(f => !f.Items.Any(i => i.ShowDetails), false);
        }
        set
        {
            foreach (XlsxPivotFieldItem item in GetField().Items)
                item.ShowDetails = !value;
        }
    }

    public XlsxPivotSortType SortType
    {
        get => GetFieldValue(f => f.SortType, XlsxPivotSortType.Default);
        set => GetField().SortType = value;
    }

    public IReadOnlyList<XlsxCellValue> SelectedValues => [];

    public IXlsxPivotFieldStyleFormats StyleFormats => new XlsxPivotTableAxisFieldStyleFormats(_pivotTable, this);

    public bool IsOnRowAxis => GetFieldValue(f => f.Axis == XlsxPivotAxis.AxisRow, _pivotTable.DataOnRows);

    public bool IsOnColumnAxis => GetFieldValue(f => f.Axis == XlsxPivotAxis.AxisCol, !_pivotTable.DataOnRows);

    public bool IsInFilterList => false;

    public int Offset => _index;

    public IXlsxPivotField SetCustomName(string value)
    {
        CustomName = value;
        return this;
    }

    public IXlsxPivotField SetSubtotalCaption(string value)
    {
        SubtotalCaption = value;
        return this;
    }

    public IXlsxPivotField AddSubtotal(XlsxSubtotalFunction value)
    {
        GetField().AddSubtotal(value);
        return this;
    }

    public IXlsxPivotField SetIncludeNewItemsInFilter(bool value)
    {
        IncludeNewItemsInFilter = value;
        return this;
    }

    public IXlsxPivotField SetLayout(XlsxPivotLayout value)
    {
        GetField().SetLayout(value);
        return this;
    }

    public IXlsxPivotField SetSubtotalsAtTop(bool value)
    {
        SubtotalsAtTop = value;
        return this;
    }

    public IXlsxPivotField SetRepeatItemLabels(bool value)
    {
        RepeatItemLabels = value;
        return this;
    }

    public IXlsxPivotField SetInsertBlankLines(bool value)
    {
        InsertBlankLines = value;
        return this;
    }

    public IXlsxPivotField SetShowBlankItems(bool value)
    {
        ShowBlankItems = value;
        return this;
    }

    public IXlsxPivotField SetInsertPageBreaks(bool value)
    {
        InsertPageBreaks = value;
        return this;
    }

    public IXlsxPivotField SetCollapsed(bool value)
    {
        Collapsed = true;
        return this;
    }

    public IXlsxPivotField SetSort(XlsxPivotSortType value)
    {
        SortType = value;
        return this;
    }

    public IXlsxPivotField AddSelectedValue(XlsxCellValue value) => this;

    public IXlsxPivotField AddSelectedValues(IEnumerable<XlsxCellValue> values) => this;

    #endregion IXlsxPivotField members

    internal XlsxPivotAxis Axis => IsOnColumnAxis ? XlsxPivotAxis.AxisCol : XlsxPivotAxis.AxisRow;

    /// <summary>
    /// Get position of the field on the axis, starting at 0.
    /// </summary>
    internal int Position
    {
        get
        {
            XlsxPivotTableAxis axis = IsOnColumnAxis ? _pivotTable.ColumnAxis : _pivotTable.RowAxis;
            var position = axis.IndexOf(_index);
            if (position == -1)
                throw new InvalidOperationException("Field is not on the axis.");

            return position;
        }
    }

    private XlsxPivotTableField GetField()
    {
        if (_index.IsDataField)
            throw new InvalidOperationException("Can't set this property on a data field.");

        return _pivotTable.PivotFields[_index];
    }

    private T GetFieldValue<T>(Func<XlsxPivotTableField, T> getter, T dataFieldValue)
    {
        if (_index.IsDataField)
            return dataFieldValue;
        XlsxPivotTableField field = _pivotTable.PivotFields[_index];
        return getter(field);
    }
}
