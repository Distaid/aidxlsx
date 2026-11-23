using System.Diagnostics;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.PivotTables.PivotValues;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.PivotTables.Areas;

/// <summary>
/// A field that describes calculation of value to display in the <see cref="XlsxPivotAreaType.Data"/>
/// area of pivot table.
/// </summary>
internal class XlsxPivotDataField : IXlsxPivotValue
{
    private const int BaseFieldDefaultValue = -1;
    private const int BaseItemPreviousValue = 1048828;
    private const int BaseItemNextValue = 1048829;
    private const int BaseItemDefaultValue = 1048832;

    private readonly XlsxPivotTable _pivotTable;

    private XlsxNumberFormat? _numberFormat;
    private int _baseField = BaseFieldDefaultValue;
    private uint _baseItem = BaseItemDefaultValue;
    private XlsxPivotCalculation _showDataAsFormat = XlsxPivotCalculation.Normal;
    private XlsxPivotSummary _subtotal = XlsxPivotSummary.Sum;

    internal XlsxPivotDataField(XlsxPivotTable pivotTable, int field)
    {
        _pivotTable = pivotTable;
        Field = field;
    }

    /// <summary>
    /// Custom name of the data field (e.g. <em>Sum of Sold</em>). Can be left empty to keep same
    /// as source name. Use <see cref="CustomName"/> to get value with fallback.
    /// </summary>
    /// <remarks>
    /// For data fields, the name is duplicated at <see cref="XlsxPivotTableField.Name"/> and here.
    /// This property has a preference.
    /// </remarks>
    internal string? DataFieldName { get; set; }

    /// <summary>
    /// Field index to <see cref="XlsxPivotTable.PivotFields"/>.
    /// </summary>
    /// <remarks>
    /// Unlike axis, this field index can't be <c>-2</c> for data fields. That field can't be in
    /// the data area.
    /// </remarks>
    internal int Field { get; }

    /// <summary>
    /// An aggregation function that calculates the value to display in the data cells of pivot area.
    /// </summary>
    public XlsxPivotSummary Subtotal
    {
        get => _subtotal;
        set => _subtotal = value;
    }

    /// <summary>
    /// A calculation takes value calculated by <see cref="Subtotal"/> aggregation and transforms
    /// it into the final value to display to the user. The calculation might need
    /// <see cref="BaseField"/> and/or <see cref="BaseItem"/>.
    /// </summary>
    public XlsxPivotCalculation ShowDataAsFormat
    {
        get => _showDataAsFormat;
        init => _showDataAsFormat = value;
    }

    /// <summary>
    /// Index to the base field (<see cref="XlsxPivotTable.PivotFields"/>) when
    /// <see cref="ShowDataAsFormat"/> needs a field for its calculation.
    /// </summary>
    public int BaseField
    {
        get => _baseField;
        init => _baseField = value;
    }

    /// <summary>
    /// Index to the base item of <see cref="BaseField"/> when <see cref="ShowDataAsFormat"/> needs
    /// an item for its calculation.
    /// </summary>
    public uint BaseItem
    {
        get => _baseItem;
        init => _baseItem = value;
    }

    /// <summary>
    /// Formatting to apply to the data field. If <see cref="XlsxPivotFormat"/> disagree, this has precedence.
    /// </summary>
    internal XlsxNumberFormat? NumberFormatValue
    {
        get => _numberFormat;
        set
        {
            _numberFormat = value is not null
                ? _pivotTable.Worksheet.Workbook.Styles.RegisterNumberFormat(value.Value)
                : null;
        }
    }

    #region IXlsxPivotValue

    public string? BaseFieldName
    {
        get
        {
            IReadOnlyList<string> sourceNames = _pivotTable.PivotCache.FieldNames;
            if (_baseField < 0 || _baseField >= sourceNames.Count)
                return null;

            return sourceNames[_baseField];
        }
        set
        {
            if (value is null)
            {
                _baseField = BaseFieldDefaultValue;
                return;
            }

            if (!_pivotTable.TryGetSourceNameFieldIndex(value, out FieldIndex index))
            {
                throw new ArgumentOutOfRangeException($"Source name '{value}' not found.");
            }

            _baseField = index;
        }
    }

    public XlsxCellValue BaseItemValue
    {
        get
        {
            var baseFieldSpecified = _baseField != BaseFieldDefaultValue;
            if (!baseFieldSpecified)
                return Blank.Value;

            var baseItemSpecified = _baseItem != BaseItemDefaultValue;
            if (!baseItemSpecified)
                return Blank.Value;

            if (_baseItem == BaseItemPreviousValue)
                return Blank.Value;

            if (_baseItem == BaseItemNextValue)
                return Blank.Value;

            XlsxPivotTableField baseField = _pivotTable.PivotFields[_baseField];
            XlsxPivotFieldItem fieldItem = baseField.Items[checked((int)BaseItem)];
            return fieldItem.GetValue() ?? Blank.Value;
        }
        set
        {
            if (_baseField == BaseItemDefaultValue)
                throw new InvalidOperationException("Base field not specified for the field.");

            XlsxPivotTableField pivotField = _pivotTable.PivotFields[_baseField];
            XlsxPivotFieldItem fieldItem = pivotField.GetOrAddItem(value);
            var itemIndex = fieldItem.ItemIndex ?? BaseFieldDefaultValue;
            _baseItem = checked((uint)itemIndex);
        }
    }

    public XlsxPivotCalculation Calculation
    {
        get => ShowDataAsFormat;
        set => _showDataAsFormat = value;
    }

    public XlsxPivotCalculationItem CalculationItem
    {
        get
        {
            return _baseItem switch
            {
                BaseItemPreviousValue => XlsxPivotCalculationItem.Previous,
                BaseItemNextValue => XlsxPivotCalculationItem.Next,
                _ => XlsxPivotCalculationItem.Value,
            };
        }
        set
        {
            switch (value)
            {
                case XlsxPivotCalculationItem.Previous:
                    _baseItem = BaseItemPreviousValue;
                    break;
                case XlsxPivotCalculationItem.Next:
                    _baseItem = BaseItemNextValue;
                    break;
                case XlsxPivotCalculationItem.Value:
                    // Calculation value should be set in tandem with the base item value.
                    // Base item other than prev/next special constants is implicitly a value.
                    if (BaseItem is BaseItemPreviousValue or BaseItemNextValue)
                    {
                        // If value is not yet set, just use unspecified value. User should
                        // set value by calling `BaseItemValue` after calling this, but Excel
                        // accepts valid base field with unspecified item without need to repair.
                        _baseItem = BaseItemDefaultValue;
                    }

                    // When base item is not a valid reference to the field.Items, Excel
                    // tries to repair the workbook, so user should always set base value.
                    break;
                default:
                    throw new UnreachableException();
            }
        }
    }

    public string CustomName
    {
        get => DataFieldName ?? _pivotTable.PivotFields[Field].Name ?? _pivotTable.PivotCache.FieldNames[Field];
        set => DataFieldName = value;
    }

    public IXlsxPivotValueFormat NumberFormat => new XlsxPivotValueFormat(this);

    public string SourceName => _pivotTable.PivotCache.FieldNames[Field];

    public XlsxPivotSummary SummaryFormula
    {
        get => Subtotal;
        set => _subtotal = value;
    }

    public IXlsxPivotValue SetBaseFieldName(string value)
    {
        BaseFieldName = value;
        return this;
    }

    public IXlsxPivotValue SetBaseItemValue(XlsxCellValue value)
    {
        BaseItemValue = value;
        return this;
    }

    public IXlsxPivotValue SetCalculation(XlsxPivotCalculation value)
    {
        Calculation = value;
        return this;
    }

    public IXlsxPivotValue SetCalculationItem(XlsxPivotCalculationItem value)
    {
        CalculationItem = value;
        return this;
    }

    public IXlsxPivotValue SetSummaryFormula(XlsxPivotSummary value)
    {
        SummaryFormula = value;
        return this;
    }

    public IXlsxPivotValueCombination ShowAsDifferenceFrom(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        SetCalculation(XlsxPivotCalculation.DifferenceFrom);
        return new XlsxPivotValueCombination(this);
    }

    public IXlsxPivotValue ShowAsIndex()
    {
        return SetCalculation(XlsxPivotCalculation.Index);
    }

    public IXlsxPivotValue ShowAsNormal()
    {
        return SetCalculation(XlsxPivotCalculation.Normal);
    }

    public IXlsxPivotValueCombination ShowAsPercentageDifferenceFrom(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        SetCalculation(XlsxPivotCalculation.PercentageDifferenceFrom);
        return new XlsxPivotValueCombination(this);
    }

    public IXlsxPivotValueCombination ShowAsPercentageFrom(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        SetCalculation(XlsxPivotCalculation.PercentageOf);
        return new XlsxPivotValueCombination(this);
    }

    public IXlsxPivotValue ShowAsPercentageOfColumn()
    {
        return SetCalculation(XlsxPivotCalculation.PercentageOfColumn);
    }

    public IXlsxPivotValue ShowAsPercentageOfRow()
    {
        return SetCalculation(XlsxPivotCalculation.PercentageOfRow);
    }

    public IXlsxPivotValue ShowAsPercentageOfTotal()
    {
        return SetCalculation(XlsxPivotCalculation.PercentageOfTotal);
    }

    public IXlsxPivotValue ShowAsRunningTotalIn(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        return SetCalculation(XlsxPivotCalculation.RunningTotal);
    }

    #endregion IXPivotValue
}
