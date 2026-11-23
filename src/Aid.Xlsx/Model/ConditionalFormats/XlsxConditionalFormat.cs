#nullable disable warnings

using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.ConditionalFormats;

internal class XlsxConditionalFormat : IXlsxDxfContainer, IXlsxConditionalFormat
{
    private readonly XlsxWorksheet _worksheet;
    private sealed class NoRangeCfComparer : IEqualityComparer<XlsxConditionalFormat>
    {
        public bool Equals(XlsxConditionalFormat? xx, XlsxConditionalFormat? yy)
        {
            if (ReferenceEquals(xx, yy)) return true;
            if (xx is null) return false;
            if (yy is null) return false;
            if (xx.GetType() != yy.GetType()) return false;

            IEnumerable<string> xxValues = xx.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
            IEnumerable<string> yyValues = yy.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
            IEnumerable<string> xxFormulas = xx.Areas.Count > 0 ? xx.Values.Values.Where(v => v != null && v.IsFormula).Select(f => ((XlsxCell)xx.Range.FirstCell()).GetFormulaR1C1(f.Value)) : null;
            IEnumerable<string> yyFormulas = yy.Areas.Count > 0 ? yy.Values.Values.Where(v => v != null && v.IsFormula).Select(f => ((XlsxCell)yy.Range.FirstCell()).GetFormulaR1C1(f.Value)) : null;
            XlsxDxfValue xStyle = xx.FormatValue;
            XlsxDxfValue yStyle = yy.FormatValue;
            return Equals(xStyle, yStyle)
                && xx.ConditionalFormatType == yy.ConditionalFormatType
                && xx.TimePeriod == yy.TimePeriod
                && xx.IconSetStyle == yy.IconSetStyle
                && xx.Operator == yy.Operator
                && xx.Bottom == yy.Bottom
                && xx.Percent == yy.Percent
                && xx.ReverseIconOrder == yy.ReverseIconOrder
                && xx.StopIfTrue == yy.StopIfTrue
                && xx.ShowIconOnly == yy.ShowIconOnly
                && xx.ShowBarOnly == yy.ShowBarOnly
                && SetEquals(xxValues, yyValues)
                && SetEquals(xxFormulas, yyFormulas)
                && Equals(xx.Colors, yy.Colors)
                && Equals(xx.ContentTypes, yy.ContentTypes)
                && Equals(xx.IconSetOperators, yy.IconSetOperators);
        }

        public int GetHashCode(XlsxConditionalFormat obj)
        {
            XlsxConditionalFormat xx = obj;
            XlsxDxfValue xStyle = obj.FormatValue;
            IEnumerable<string> xValues = xx.Values.Values.Where(v => !v.IsFormula).Select(v => v.Value);
            if (obj.Areas.Count > 0)
                xValues = xValues
                .Union(xx.Values.Values.Where(v => v.IsFormula).Select(f => ((XlsxCell)obj.Range.FirstCell()).GetFormulaR1C1(f.Value)));

            unchecked
            {
                var hashCode = xStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ xValues.GetHashCode();
                hashCode = (hashCode * 397) ^ (xx.Colors != null ? xx.Colors.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (xx.ContentTypes != null ? xx.ContentTypes.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (xx.IconSetOperators != null ? xx.IconSetOperators.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (int)xx.ConditionalFormatType;
                hashCode = (hashCode * 397) ^ (int)xx.TimePeriod;
                hashCode = (hashCode * 397) ^ (int)xx.IconSetStyle;
                hashCode = (hashCode * 397) ^ (int)xx.Operator;
                hashCode = (hashCode * 397) ^ xx.Bottom.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.Percent.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ReverseIconOrder.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ShowIconOnly.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ShowBarOnly.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.StopIfTrue.GetHashCode();
                return hashCode;
            }
        }

        private static bool SetEquals<T>(IEnumerable<T> first, IEnumerable<T> second)
        {
            return new HashSet<T>(second, EqualityComparer<T>.Default)
                .SetEquals(first);
        }

        private static bool Equals<TValue>(Dictionary<int, TValue> x, Dictionary<int, TValue> y)
        {
            if (x.Count != y.Count)
                return false;
            if (x.Keys.Except(y.Keys).Any())
                return false;
            if (y.Keys.Except(x.Keys).Any())
                return false;
            EqualityComparer<TValue> valueComparer = EqualityComparer<TValue>.Default;
            foreach (KeyValuePair<int, TValue> pair in x)
                if (!valueComparer.Equals(pair.Value, y[pair.Key]))
                    return false;

            return true;
        }
    }

    internal static IEqualityComparer<XlsxConditionalFormat> NoRangeComparer { get; } = new NoRangeCfComparer();

    #region Constructors

    internal XlsxConditionalFormat(XlsxWorksheet worksheet, XlsxAreaList areaList)
    {
        _worksheet = worksheet;
        Id = Guid.NewGuid();
        Areas = areaList;
        Values = [];
        Colors = [];
        ContentTypes = [];
        IconSetOperators = [];
    }

    /// <summary>
    /// Copy ctor.
    /// </summary>
    internal XlsxConditionalFormat(XlsxWorksheet worksheet, XlsxConditionalFormat other, XlsxAreaList areaList)
        : this(worksheet, areaList)
    {
        XlsxDxfValue otherDxf = other.FormatValue;
        FormatValue = otherDxf is not null
            ? _worksheet.Workbook.Styles.GetRegisteredDxFormat(otherDxf, static x => x)
            : null;
        ConditionalFormatType = other.ConditionalFormatType;
        TimePeriod = other.TimePeriod;
        IconSetStyle = other.IconSetStyle;
        Operator = other.Operator;
        Bottom = other.Bottom;
        Percent = other.Percent;
        ReverseIconOrder = other.ReverseIconOrder;
        ShowIconOnly = other.ShowIconOnly;
        ShowBarOnly = other.ShowBarOnly;
        StopIfTrue = other.StopIfTrue;

        Point sourceAnchor = other.Areas[0].FirstPoint;
        Point targetAnchor = Areas[0].FirstPoint;
        foreach ((int key, XlsxFormula originalValue) in other.Values)
            Values.Add(key, originalValue.GetAdjustedCopy(sourceAnchor, targetAnchor));

        Colors = other.Colors.CopyDictionary();
        ContentTypes = other.ContentTypes.CopyDictionary();
        IconSetOperators = other.IconSetOperators.CopyDictionary();
    }

    #endregion Constructors

    public Guid Id { get; internal set; }

    /// <summary>
    /// Priority of formatting rule. Lower values have higher priority than higher values.
    /// Minimum value is 1. It is basically used for ordering of CF during saving.
    /// </summary>
    internal int Priority { get; set; }

    public XlsxDxfValue? FormatValue { get; set; }

    internal XlsxDxFormat Format => new(_worksheet.Workbook.Styles, this);

    public IXlsxStyle Style
    {
        get => Format;
        set => Format.SetStyle(value);
    }

    public XlsxDictionary<XlsxFormula> Values { get; }

    public XlsxDictionary<XlsxColor> Colors { get; }

    public XlsxDictionary<XlsxCFContentType> ContentTypes { get; }

    public XlsxDictionary<XlsxCFIconSetOperator> IconSetOperators { get; }

    public IXlsxRange Range
    {
        get => _worksheet.Range(Areas[0]);
        set => Areas = XlsxAreaList.FromRange(_worksheet, value);
    }

    public IEnumerable<IXlsxRange> Ranges
    {
        get
        {
            var ranges = new XlsxRanges(_worksheet);
            foreach (Area area in Areas)
                ranges.Add(_worksheet.Range(area));

            return ranges;
        }
        set => Areas = XlsxAreaList.FromRanges(_worksheet, value);
    }

    public XlsxConditionalFormatType ConditionalFormatType { get; set; }

    public XlsxTimePeriod TimePeriod { get; set; }

    public XlsxIconSetStyle IconSetStyle { get; set; }

    public XlsxCFOperator Operator { get; set; }

    public bool Bottom { get; set; }

    public bool Percent { get; set; }

    public bool ReverseIconOrder { get; set; }

    public bool ShowIconOnly { get; set; }

    public bool ShowBarOnly { get; set; }

    public bool StopIfTrue { get; set; }

    internal XlsxAreaList Areas { get; set; }

    public IXlsxConditionalFormat SetStopIfTrue()
    {
        return SetStopIfTrue(true);
    }

    public IXlsxConditionalFormat SetStopIfTrue(bool value)
    {
        this.StopIfTrue = value;
        return this;
    }

    public IXlsxConditionalFormat CopyTo(IXlsxWorksheet targetSheet)
    {
        if (targetSheet == Range?.Worksheet)
            throw new InvalidOperationException("Cannot copy conditional format to the worksheet it already belongs to.");
        var newCf = new XlsxConditionalFormat((XlsxWorksheet)targetSheet, this, Areas);
        targetSheet.ConditionalFormats.Add(newCf);
        return newCf;
    }

    public IXlsxStyle WhenIsBlank()
    {
        ConditionalFormatType = XlsxConditionalFormatType.IsBlank;
        return Style;
    }

    public IXlsxStyle WhenNotBlank()
    {
        ConditionalFormatType = XlsxConditionalFormatType.NotBlank;
        return Style;
    }

    public IXlsxStyle WhenIsError()
    {
        ConditionalFormatType = XlsxConditionalFormatType.IsError;
        return Style;
    }

    public IXlsxStyle WhenNotError()
    {
        ConditionalFormatType = XlsxConditionalFormatType.NotError;
        return Style;
    }

    public IXlsxStyle WhenDateIs(XlsxTimePeriod timePeriod)
    {
        TimePeriod = timePeriod;
        ConditionalFormatType = XlsxConditionalFormatType.TimePeriod;
        return Style;
    }

    public IXlsxStyle WhenContains(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        ConditionalFormatType = XlsxConditionalFormatType.ContainsText;
        Operator = XlsxCFOperator.Contains;
        return Style;
    }

    public IXlsxStyle WhenNotContains(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        ConditionalFormatType = XlsxConditionalFormatType.NotContainsText;
        Operator = XlsxCFOperator.NotContains;
        return Style;
    }

    public IXlsxStyle WhenStartsWith(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        ConditionalFormatType = XlsxConditionalFormatType.StartsWith;
        Operator = XlsxCFOperator.StartsWith;
        return Style;
    }

    public IXlsxStyle WhenEndsWith(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        ConditionalFormatType = XlsxConditionalFormatType.EndsWith;
        Operator = XlsxCFOperator.EndsWith;
        return Style;
    }

    public IXlsxStyle WhenEquals(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        Operator = XlsxCFOperator.Equal;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenNotEquals(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        Operator = XlsxCFOperator.NotEqual;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenGreaterThan(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        Operator = XlsxCFOperator.GreaterThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenLessThan(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        Operator = XlsxCFOperator.LessThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenEqualOrGreaterThan(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        Operator = XlsxCFOperator.EqualOrGreaterThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenEqualOrLessThan(string value)
    {
        Values.Initialize(new XlsxFormula { Value = value });
        Operator = XlsxCFOperator.EqualOrLessThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenBetween(string minValue, string maxValue)
    {
        Values.Initialize(new XlsxFormula { Value = minValue });
        Values.Add(new XlsxFormula { Value = maxValue });
        Operator = XlsxCFOperator.Between;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenNotBetween(string minValue, string maxValue)
    {
        Values.Initialize(new XlsxFormula { Value = minValue });
        Values.Add(new XlsxFormula { Value = maxValue });
        Operator = XlsxCFOperator.NotBetween;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenEquals(double value)
    {
        Values.Initialize(new XlsxFormula(value));
        Operator = XlsxCFOperator.Equal;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenNotEquals(double value)
    {
        Values.Initialize(new XlsxFormula(value));
        Operator = XlsxCFOperator.NotEqual;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenGreaterThan(double value)
    {
        Values.Initialize(new XlsxFormula(value));
        Operator = XlsxCFOperator.GreaterThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenLessThan(double value)
    {
        Values.Initialize(new XlsxFormula(value));
        Operator = XlsxCFOperator.LessThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenEqualOrGreaterThan(double value)
    {
        Values.Initialize(new XlsxFormula(value));
        Operator = XlsxCFOperator.EqualOrGreaterThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenEqualOrLessThan(double value)
    {
        Values.Initialize(new XlsxFormula(value));
        Operator = XlsxCFOperator.EqualOrLessThan;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenBetween(double minValue, double maxValue)
    {
        Values.Initialize(new XlsxFormula(minValue));
        Values.Add(new XlsxFormula(maxValue));
        Operator = XlsxCFOperator.Between;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenNotBetween(double minValue, double maxValue)
    {
        Values.Initialize(new XlsxFormula(minValue));
        Values.Add(new XlsxFormula(maxValue));
        Operator = XlsxCFOperator.NotBetween;
        ConditionalFormatType = XlsxConditionalFormatType.CellIs;
        return Style;
    }

    public IXlsxStyle WhenIsDuplicate()
    {
        ConditionalFormatType = XlsxConditionalFormatType.IsDuplicate;
        return Style;
    }

    public IXlsxStyle WhenIsUnique()
    {
        ConditionalFormatType = XlsxConditionalFormatType.IsUnique;
        return Style;
    }

    public IXlsxStyle WhenIsTrue(string formula)
    {
        string f = formula.TrimStart()[0] == '=' ? formula : "=" + formula;
        Values.Initialize(new XlsxFormula { Value = f });
        ConditionalFormatType = XlsxConditionalFormatType.Expression;
        return Style;
    }

    public IXlsxStyle WhenIsTop(int value, XlsxTopBottomType topBottomType = XlsxTopBottomType.Items)
    {
        Values.Initialize(new XlsxFormula(value));
        Percent = topBottomType == XlsxTopBottomType.Percent;
        ConditionalFormatType = XlsxConditionalFormatType.Top10;
        Bottom = false;
        return Style;
    }

    public IXlsxStyle WhenIsBottom(int value, XlsxTopBottomType topBottomType = XlsxTopBottomType.Items)
    {
        Values.Initialize(new XlsxFormula(value));
        Percent = topBottomType == XlsxTopBottomType.Percent;
        ConditionalFormatType = XlsxConditionalFormatType.Top10;
        Bottom = true;
        return Style;
    }

    public IXlsxCFColorScaleMin ColorScale()
    {
        ConditionalFormatType = XlsxConditionalFormatType.ColorScale;
        return new XlsxCFColorScaleMin(this);
    }

    public IXlsxCFDataBarMin DataBar(XlsxColor color, bool showBarOnly = false)
    {
        Colors.Initialize(color);
        ShowBarOnly = showBarOnly;
        ConditionalFormatType = XlsxConditionalFormatType.DataBar;
        return new XlsxCFDataBarMin(this);
    }

    public IXlsxCFDataBarMin DataBar(XlsxColor positiveColor, XlsxColor negativeColor, bool showBarOnly = false)
    {
        Colors.Initialize(positiveColor);
        Colors.Add(negativeColor);
        ShowBarOnly = showBarOnly;
        ConditionalFormatType = XlsxConditionalFormatType.DataBar;
        return new XlsxCFDataBarMin(this);
    }

    public IXlsxCFIconSet IconSet(XlsxIconSetStyle iconSetStyle, bool reverseIconOrder = false, bool showIconOnly = false)
    {
        IconSetOperators.Clear();
        Values.Clear();
        ContentTypes.Clear();
        ConditionalFormatType = XlsxConditionalFormatType.IconSet;
        IconSetStyle = iconSetStyle;
        ReverseIconOrder = reverseIconOrder;
        ShowIconOnly = showIconOnly;
        return new XlsxCFIconSet(this);
    }
}
