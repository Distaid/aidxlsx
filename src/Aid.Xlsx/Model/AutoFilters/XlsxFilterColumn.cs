using System.Collections;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.AutoFilters;

internal class XlsxFilterColumn(XlsxAutoFilter autoFilter, Int32 column) : IXlsxFilterColumn, IXlsxFilteredColumn, IEnumerable<XlsxFilter>
{
    private readonly XlsxAutoFilter _autoFilter = autoFilter;
    private readonly Int32 _column = column;
    private readonly List<XlsxFilter> _filters = [];

    #region IXlsxFilterColumn Members

    public void Clear(bool reapply)
    {
        _filters.Clear();
        FilterType = XlsxFilterType.None;
        if (reapply)
            _autoFilter.Reapply();
    }

    public IXlsxFilteredColumn AddFilter(XlsxCellValue value, bool reapply)
    {
        SwitchFilter(XlsxFilterType.Regular);
        AddFilter(XlsxFilter.CreateRegularFilter(value.ToString()), reapply);
        return this;
    }

    public IXlsxFilteredColumn AddDateGroupFilter(DateTime date, XlsxDateTimeGrouping dateTimeGrouping, bool reapply)
    {
        SwitchFilter(XlsxFilterType.Regular);
        AddFilter(XlsxFilter.CreateDateGroupFilter(date, dateTimeGrouping), reapply);
        return this;
    }

    public void Top(Int32 value, XlsxTopBottomType type, bool reapply)
    {
        SetTopBottom(value, type, takeTop: true, reapply);
    }

    public void Bottom(Int32 value, XlsxTopBottomType type, bool reapply)
    {
        SetTopBottom(value, type, takeTop: false, reapply);
    }

    public void AboveAverage(bool reapply)
    {
        SetAverage(aboveAverage: true, reapply);
    }

    public void BelowAverage(bool reapply)
    {
        SetAverage(aboveAverage: false, reapply);
    }

    public IXlsxFilterConnector EqualTo(XlsxCellValue value, Boolean reapply)
    {
        return AddCustomFilter(value.ToString(), true, reapply);
    }

    public IXlsxFilterConnector NotEqualTo(XlsxCellValue value, Boolean reapply)
    {
        return AddCustomFilter(value.ToString(), false, reapply);
    }

    public IXlsxFilterConnector GreaterThan(XlsxCellValue value, Boolean reapply)
    {
        return AddCustomFilter(value, XlsxFilterOperator.GreaterThan, reapply);
    }

    public IXlsxFilterConnector LessThan(XlsxCellValue value, Boolean reapply)
    {
        return AddCustomFilter(value, XlsxFilterOperator.LessThan, reapply);
    }

    public IXlsxFilterConnector EqualOrGreaterThan(XlsxCellValue value, Boolean reapply)
    {
        return AddCustomFilter(value, XlsxFilterOperator.EqualOrGreaterThan, reapply);
    }

    public IXlsxFilterConnector EqualOrLessThan(XlsxCellValue value, Boolean reapply)
    {
        return AddCustomFilter(value, XlsxFilterOperator.EqualOrLessThan, reapply);
    }

    public void Between(XlsxCellValue minValue, XlsxCellValue maxValue, Boolean reapply)
    {
        EqualOrGreaterThan(minValue, false).And.EqualOrLessThan(maxValue, reapply);
    }

    public void NotBetween(XlsxCellValue minValue, XlsxCellValue maxValue, Boolean reapply)
    {
        LessThan(minValue, false).Or.GreaterThan(maxValue, reapply);
    }

    public IXlsxFilterConnector BeginsWith(String value, Boolean reapply)
    {
        return AddCustomFilter(value + "*", true, reapply);
    }

    public IXlsxFilterConnector NotBeginsWith(String value, Boolean reapply)
    {
        return AddCustomFilter(value + "*", false, reapply);
    }

    public IXlsxFilterConnector EndsWith(String value, Boolean reapply)
    {
        return AddCustomFilter("*" + value, true, reapply);
    }

    public IXlsxFilterConnector NotEndsWith(String value, Boolean reapply)
    {
        return AddCustomFilter("*" + value, false, reapply);
    }

    public IXlsxFilterConnector Contains(String value, Boolean reapply)
    {
        return AddCustomFilter("*" + value + "*", true, reapply);
    }

    public IXlsxFilterConnector NotContains(String value, Boolean reapply)
    {
        return AddCustomFilter("*" + value + "*", false, reapply);
    }

    public XlsxFilterType FilterType { get; set; }

    public Int32 TopBottomValue { get; set; }
    public XlsxTopBottomType TopBottomType { get; set; }
    public XlsxTopBottomPart TopBottomPart { get; set; }

    public XlsxFilterDynamicType DynamicType { get; set; }

    /// <summary>
    /// Basically average for dynamic filters. Value is refreshed during filter reapply.
    /// </summary>
    public Double DynamicValue { get; set; } = double.NaN;

    #endregion IXlsxFilterColumn Members

    /// <summary>
    /// A filter value used by top/bottom filter to compare with cell value.
    /// </summary>
    internal double TopBottomFilterValue { get; private set; } = double.NaN;

    public IEnumerator<XlsxFilter> GetEnumerator() => _filters.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private void SetTopBottom(Int32 percentOrItemCount, XlsxTopBottomType type, Boolean takeTop, Boolean reapply)
    {
        if (percentOrItemCount is < 1 or > 500)
            throw new ArgumentOutOfRangeException(nameof(percentOrItemCount), "Value must be between 1 and 500.");

        ResetFilter(XlsxFilterType.TopBottom);
        TopBottomValue = percentOrItemCount;
        TopBottomType = type;
        TopBottomPart = takeTop ? XlsxTopBottomPart.Top : XlsxTopBottomPart.Bottom;

        AddFilter(XlsxFilter.CreateTopBottom(takeTop, percentOrItemCount), reapply);
    }

    private double GetTopBottomFilterValue(XlsxTopBottomType type, int value, bool takeTop)
    {
        var column = _autoFilter.Range.Column(_column);
        var subColumn = column.Column(2, column.CellCount());
        var columnNumbers = subColumn.CellsUsed(c => c.CachedValue.IsUnifiedNumber).Select(c => c.CachedValue.GetUnifiedNumber());
        var comparer = takeTop
            ? Comparer<double>.Create((x, y) => -x.CompareTo(y))
            : Comparer<double>.Create((x, y) => x.CompareTo(y));

        switch (type)
        {
            case XlsxTopBottomType.Items:
                var itemCount = value;
                return columnNumbers.OrderBy(d => d, comparer).Take(itemCount).DefaultIfEmpty(double.NaN).LastOrDefault();
            case XlsxTopBottomType.Percent:
                var percent = value;
                var materializedNumbers = columnNumbers.ToArray();

                // Ceiling, so there is always at least one item.
                var itemCountByPercents = (int)Math.Ceiling(materializedNumbers.Length * (double)percent / 100);
                return materializedNumbers.OrderBy(d => d, comparer).Take(itemCountByPercents).DefaultIfEmpty(Double.NaN).LastOrDefault();
            default:
                throw new NotSupportedException();
        }
    }

    private void SetAverage(Boolean aboveAverage, Boolean reapply)
    {
        ResetFilter(XlsxFilterType.Dynamic);
        DynamicType = aboveAverage
            ? XlsxFilterDynamicType.AboveAverage
            : XlsxFilterDynamicType.BelowAverage;

        // `Average` is recalculated during reapply, so no need to calculate it twice.
        DynamicValue = reapply ? double.NaN : GetAverageFilterValue();
        AddFilter(XlsxFilter.CreateAverage(DynamicValue, aboveAverage), reapply);
    }

    private double GetAverageFilterValue()
    {
        var column = _autoFilter.Range.Column(_column);
        var subColumn = column.Column(2, column.CellCount());
        return subColumn.CellsUsed(c => c.CachedValue.IsUnifiedNumber)
            .Select(c => c.CachedValue.GetUnifiedNumber())
            .DefaultIfEmpty(Double.NaN)
            .Average();
    }

    private IXlsxFilterConnector AddCustomFilter(XlsxCellValue value, XlsxFilterOperator op, Boolean reapply)
    {
        ResetFilter(XlsxFilterType.Custom);
        AddFilter(XlsxFilter.CreateCustomFilter(value, op, XlsxConnector.Or), reapply);
        return new XlsxFilterConnector(this);
    }

    private IXlsxFilterConnector AddCustomFilter(string pattern, bool match, bool reapply)
    {
        ResetFilter(XlsxFilterType.Custom);
        AddFilter(XlsxFilter.CreateCustomPatternFilter(pattern, match, XlsxConnector.Or), reapply);
        return new XlsxFilterConnector(this);
    }

    private void ResetFilter(XlsxFilterType type)
    {
        Clear(false);
        _autoFilter.IsEnabled = true;
        FilterType = type;
    }

    private void SwitchFilter(XlsxFilterType type)
    {
        _autoFilter.IsEnabled = true;
        if (FilterType == type)
            return;

        Clear(false);
        FilterType = type;
    }

    internal void AddFilter(XlsxFilter filter, bool reapply = false)
    {
        var maxFilters = FilterType switch
        {
            XlsxFilterType.None => 0,
            XlsxFilterType.Regular => int.MaxValue,
            XlsxFilterType.Custom => 2,
            XlsxFilterType.TopBottom => 1,
            XlsxFilterType.Dynamic => 1,
            _ => throw new NotSupportedException()
        };
        if (_filters.Count >= maxFilters)
            throw new InvalidOperationException($"{FilterType} filter can have max {maxFilters} conditions.");

        _filters.Add(filter);
        if (reapply)
            _autoFilter.Reapply();
    }

    internal void Refresh()
    {
        if (FilterType == XlsxFilterType.Dynamic)
        {
            // Update average value of a filter, so it is saved correctly and filter uses
            // correct value, even is cell values changed and avg was stale.
            DynamicValue = GetAverageFilterValue();
            _filters[0].Value = DynamicValue;
        }

        if (FilterType == XlsxFilterType.TopBottom)
        {
            var takeTop = TopBottomPart == XlsxTopBottomPart.Top;
            TopBottomFilterValue = GetTopBottomFilterValue(TopBottomType, TopBottomValue, takeTop);
        }
    }

    internal bool Check(IXlsxCell cell)
    {
        if (_filters.Count == 0)
            return true;

        if (_filters.Count == 1)
            return _filters[0].Condition(cell, this);

        // All filter conditions are connected by a single type of logical condition. Regular
        // filters use 'Or', custom has up to two clauses connected by 'And'/'Or' and rest is
        // single clause.
        var connector = _filters[1].Connector;
        return connector switch
        {
            XlsxConnector.And => _filters.All(filter => filter.Condition(cell, this)),
            XlsxConnector.Or => _filters.Any(filter => filter.Condition(cell, this)),
            _ => throw new NotSupportedException(),
        };
    }
}
