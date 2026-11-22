#nullable disable

using Aid.Xlsx.AutoFilters;
using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.ConditionalFormats
{
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

        IXlsxStyle WhenContains(String value);

        IXlsxStyle WhenNotContains(String value);

        IXlsxStyle WhenStartsWith(String value);

        IXlsxStyle WhenEndsWith(String value);

        IXlsxStyle WhenEquals(String value);

        IXlsxStyle WhenNotEquals(String value);

        IXlsxStyle WhenGreaterThan(String value);

        IXlsxStyle WhenLessThan(String value);

        IXlsxStyle WhenEqualOrGreaterThan(String value);

        IXlsxStyle WhenEqualOrLessThan(String value);

        IXlsxStyle WhenBetween(String minValue, String maxValue);

        IXlsxStyle WhenNotBetween(String minValue, String maxValue);

        IXlsxStyle WhenEquals(Double value);

        IXlsxStyle WhenNotEquals(Double value);

        IXlsxStyle WhenGreaterThan(Double value);

        IXlsxStyle WhenLessThan(Double value);

        IXlsxStyle WhenEqualOrGreaterThan(Double value);

        IXlsxStyle WhenEqualOrLessThan(Double value);

        IXlsxStyle WhenBetween(Double minValue, Double maxValue);

        IXlsxStyle WhenNotBetween(Double minValue, Double maxValue);

        IXlsxStyle WhenIsDuplicate();

        IXlsxStyle WhenIsUnique();

        IXlsxStyle WhenIsTrue(String formula);

        IXlsxStyle WhenIsTop(Int32 value, XlsxTopBottomType topBottomType = XlsxTopBottomType.Items);

        IXlsxStyle WhenIsBottom(Int32 value, XlsxTopBottomType topBottomType);

        IXlsxCFColorScaleMin ColorScale();

        IXlsxCFDataBarMin DataBar(XlsxColor color, Boolean showBarOnly = false);

        IXlsxCFDataBarMin DataBar(XlsxColor positiveColor, XlsxColor negativeColor, Boolean showBarOnly = false);

        IXlsxCFIconSet IconSet(XlsxIconSetStyle iconSetStyle, Boolean reverseIconOrder = false, Boolean showIconOnly = false);

        XlsxConditionalFormatType ConditionalFormatType { get; }

        XlsxIconSetStyle IconSetStyle { get; }

        XlsxTimePeriod TimePeriod { get; }

        Boolean ReverseIconOrder { get; }

        Boolean ShowIconOnly { get; }

        Boolean ShowBarOnly { get; }

        Boolean StopIfTrue { get; }

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

        Boolean Bottom { get; }

        Boolean Percent { get; }

        IXlsxConditionalFormat SetStopIfTrue();

        IXlsxConditionalFormat SetStopIfTrue(Boolean value);

        IXlsxConditionalFormat CopyTo(IXlsxWorksheet targetSheet);
    }
}
