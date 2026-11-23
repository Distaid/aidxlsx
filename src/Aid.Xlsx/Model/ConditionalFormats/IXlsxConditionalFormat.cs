#nullable disable

using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.ConditionalFormats;

public enum XlsxTimePeriod
{
    Yesterday,
    Today,
    Tomorrow,
    InTheLast7Days,
    LastWeek,
    ThisWeek,
    NextWeek,
    LastMonth,
    ThisMonth,
    NextMonth
}

public enum XlsxIconSetStyle
{
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters
}

public enum XlsxConditionalFormatType
{
    Expression,
    CellIs,
    ColorScale,
    DataBar,
    IconSet,
    Top10,
    IsUnique,
    IsDuplicate,
    ContainsText,
    NotContainsText,
    StartsWith,
    EndsWith,
    IsBlank,
    NotBlank,
    IsError,
    NotError,
    TimePeriod,
    AboveAverage
}

public enum XlsxCFOperator { Equal, NotEqual, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan, Between, NotBetween, Contains, NotContains, StartsWith, EndsWith }

public interface IXlsxConditionalFormat
{
    IXlsxStyle Style { get; set; }

    IXlsxStyle WhenIsBlank();

    IXlsxStyle WhenNotBlank();

    IXlsxStyle WhenIsError();

    IXlsxStyle WhenNotError();

    IXlsxStyle WhenDateIs(XlsxTimePeriod timePeriod);

    IXlsxStyle WhenContains(string value);

    IXlsxStyle WhenNotContains(string value);

    IXlsxStyle WhenStartsWith(string value);

    IXlsxStyle WhenEndsWith(string value);

    IXlsxStyle WhenEquals(string value);

    IXlsxStyle WhenNotEquals(string value);

    IXlsxStyle WhenGreaterThan(string value);

    IXlsxStyle WhenLessThan(string value);

    IXlsxStyle WhenEqualOrGreaterThan(string value);

    IXlsxStyle WhenEqualOrLessThan(string value);

    IXlsxStyle WhenBetween(string minValue, string maxValue);

    IXlsxStyle WhenNotBetween(string minValue, string maxValue);

    IXlsxStyle WhenEquals(double value);

    IXlsxStyle WhenNotEquals(double value);

    IXlsxStyle WhenGreaterThan(double value);

    IXlsxStyle WhenLessThan(double value);

    IXlsxStyle WhenEqualOrGreaterThan(double value);

    IXlsxStyle WhenEqualOrLessThan(double value);

    IXlsxStyle WhenBetween(double minValue, double maxValue);

    IXlsxStyle WhenNotBetween(double minValue, double maxValue);

    IXlsxStyle WhenIsDuplicate();

    IXlsxStyle WhenIsUnique();

    IXlsxStyle WhenIsTrue(string formula);

    IXlsxStyle WhenIsTop(int value, XlsxTopBottomType topBottomType = XlsxTopBottomType.Items);

    IXlsxStyle WhenIsBottom(int value, XlsxTopBottomType topBottomType);

    IXlsxCFColorScaleMin ColorScale();

    IXlsxCFDataBarMin DataBar(XlsxColor color, bool showBarOnly = false);

    IXlsxCFDataBarMin DataBar(XlsxColor positiveColor, XlsxColor negativeColor, bool showBarOnly = false);

    IXlsxCFIconSet IconSet(XlsxIconSetStyle iconSetStyle, bool reverseIconOrder = false, bool showIconOnly = false);

    XlsxConditionalFormatType ConditionalFormatType { get; }

    XlsxIconSetStyle IconSetStyle { get; }

    XlsxTimePeriod TimePeriod { get; }

    bool ReverseIconOrder { get; }

    bool ShowIconOnly { get; }

    bool ShowBarOnly { get; }

    bool StopIfTrue { get; }

    /// <summary>
    /// The first of the <see cref="Ranges"/>.
    /// </summary>
    IXlsxRange Range { get; set; }

    /// <summary>
    /// Get or set all ranges the conditional format applies to.
    /// </summary>
    /// <exception cref="ArgumentException">If sequence contains no elements or the range is from different worksheet.</exception>
    IEnumerable<IXlsxRange> Ranges { get; set; }

    XlsxDictionary<XlsxFormula> Values { get; }

    XlsxDictionary<XlsxColor> Colors { get; }

    XlsxDictionary<XlsxCFContentType> ContentTypes { get; }

    XlsxDictionary<XlsxCFIconSetOperator> IconSetOperators { get; }

    XlsxCFOperator Operator { get; }

    bool Bottom { get; }

    bool Percent { get; }

    IXlsxConditionalFormat SetStopIfTrue();

    IXlsxConditionalFormat SetStopIfTrue(bool value);

    IXlsxConditionalFormat CopyTo(IXlsxWorksheet targetSheet);
}
