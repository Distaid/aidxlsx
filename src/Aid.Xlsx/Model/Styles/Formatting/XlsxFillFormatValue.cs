using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles.Formatting;

/// <summary>
/// A fill formatting. Fill can have at most one pattern that specifies how to fill a cell.
/// </summary>
internal record XlsxFillFormatValue
{
    internal static readonly XlsxFillFormatValue Empty = new(null, null, null);

    internal static readonly XlsxFillFormatValue None = new(new XlsxPatternFill
    {
        PatternType = XlsxFillPatternValues.None,
        BackgroundColor = XlsxColor.Automatic,
        PatternColor = XlsxColor.Automatic
    });

    internal static readonly XlsxFillFormatValue Gray125 = new(new XlsxPatternFill
    {
        PatternType = XlsxFillPatternValues.Gray125,
        BackgroundColor = XlsxColor.Automatic,
        PatternColor = XlsxColor.Automatic
    });

    public XlsxFillFormatValue(XlsxPatternFill pattern)
        : this(pattern, null, null)
    {
    }

    public XlsxFillFormatValue(XlsxLinearGradientFill linearGradient)
        : this(null, linearGradient, null)
    {
    }

    public XlsxFillFormatValue(XlsxPathGradientFill pathGradient)
        : this(null, null, pathGradient)
    {
    }

    private XlsxFillFormatValue(XlsxPatternFill? pattern, XlsxLinearGradientFill? linearGradient, XlsxPathGradientFill? pathGradient)
    {
        Pattern = pattern;
        LinearGradient = linearGradient;
        PathGradient = pathGradient;
    }

    public XlsxPatternFill? Pattern { get; }

    public XlsxLinearGradientFill? LinearGradient { get; }

    public XlsxPathGradientFill? PathGradient { get; }
}
