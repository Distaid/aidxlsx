#nullable disable
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

namespace Aid.Xlsx.Model.PivotTables.Areas;

/// <summary>
/// Fluent API for filter fields of a <see cref="XlsxPivotTable"/>. This class shouldn't contain any
/// state, only logic to change state per API.
/// </summary>
internal class XlsxPivotTablePageField : IXlsxPivotField
{
    private readonly XlsxPivotTable _pivotTable;
    private readonly XlsxPivotPageField _filterField;

    internal XlsxPivotTablePageField(XlsxPivotTable pivotTable, XlsxPivotPageField filterField)
    {
        _pivotTable = pivotTable;
        _filterField = filterField;
    }

    public string SourceName => _pivotTable.PivotCache.FieldNames[_filterField.Field];

    public string CustomName
    {
        get => GetField().Name;
        set => GetField().Name = value;
    }

    public string SubtotalCaption
    {
        get => GetField().SubtotalCaption;
        set => GetField().SubtotalCaption = value;
    }

    public IReadOnlyCollection<XlsxSubtotalFunction> Subtotals => GetField().Subtotals;

    public bool IncludeNewItemsInFilter
    {
        get => GetField().IncludeNewItemsInFilter;
        set => GetField().IncludeNewItemsInFilter = value;
    }

    public bool Outline
    {
        get => GetField().Outline;
        set => GetField().Outline = value;
    }

    public bool Compact
    {
        get => GetField().Compact;
        set => GetField().Compact = value;
    }

    public bool? SubtotalsAtTop
    {
        get => GetField().SubtotalTop;
        set => GetField().SubtotalTop = value ?? true;
    }

    public bool RepeatItemLabels
    {
        get => GetField().RepeatItemLabels;
        set => GetField().RepeatItemLabels = value;
    }

    public bool InsertBlankLines
    {
        get => GetField().InsertBlankRow;
        set => GetField().InsertBlankRow = value;
    }

    public bool ShowBlankItems
    {
        get => GetField().ShowAll;
        set => GetField().ShowAll = value;
    }

    public bool InsertPageBreaks
    {
        get => GetField().InsertPageBreak;
        set => GetField().InsertPageBreak = value;
    }

    public bool Collapsed
    {
        get => GetField().Collapsed;
        set => GetField().Collapsed = value;
    }

    public XlsxPivotSortType SortType
    {
        get => GetField().SortType;
        set => GetField().SortType = value;
    }

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
        Collapsed = value;
        return this;
    }

    public IXlsxPivotField SetSort(XlsxPivotSortType value)
    {
        SortType = value;
        return this;
    }

    public IReadOnlyList<XlsxCellValue> SelectedValues
    {
        get
        {
            var shownItems = GetField().Items.Where(i => !i.Hidden);
            var selectedValues = new List<XlsxCellValue>();
            foreach (var selectedItem in shownItems)
            {
                var selectedValue = selectedItem.GetValue();
                if (selectedValue is not null)
                    selectedValues.Add(selectedValue.Value);
            }

            return selectedValues;
        }
    }

    public IXlsxPivotField AddSelectedValue(XlsxCellValue value)
    {
        // Try to keep the original behavior - it always allows multiple selected items for added values.
        // But it's complete kludge with no sane semantic that will be nuked ASAP.
        var pivotField = GetField();

        var nothingSelected = _filterField.ItemIndex is null && !pivotField.MultipleItemSelectionAllowed;
        if (nothingSelected)
        {
            var fieldItem = pivotField.GetOrAddItem(value);
            _filterField.ItemIndex = fieldItem.ItemIndex;
            return this;
        }

        var oneItemSelected = _filterField.ItemIndex is not null && !pivotField.MultipleItemSelectionAllowed;
        if (oneItemSelected)
        {
            // Switch to multiple
            pivotField.MultipleItemSelectionAllowed = true;
            foreach (var item in pivotField.Items.Where(x => x.ItemType == XlsxPivotItemType.Data))
                item.Hidden = true;

            var selectedItem = pivotField.Items.Single(i => i.ItemIndex == _filterField.ItemIndex);
            selectedItem.Hidden = false;
            _filterField.ItemIndex = null;
            var fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
        else
        {
            // Add another item to selected item filters.
            var fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
    }

    public IXlsxPivotField AddSelectedValues(IEnumerable<XlsxCellValue> values)
    {
        foreach (var value in values)
            AddSelectedValue(value);

        return this;
    }

    public IXlsxPivotFieldStyleFormats StyleFormats => throw new NotImplementedException("Styles for filter fields are not yet implemented.");
    public bool IsOnRowAxis => false;
    public bool IsOnColumnAxis => false;
    public bool IsInFilterList => true;
    public int Offset => _filterField.Field;

    private XlsxPivotTableField GetField()
    {
        return _pivotTable.PivotFields[_filterField.Field];
    }
}
