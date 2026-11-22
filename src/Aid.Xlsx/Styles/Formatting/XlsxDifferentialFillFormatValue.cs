namespace Aid.Xlsx.Styles.Formatting;

internal record XlsxDifferentialFillValue
{
    internal static XlsxDifferentialFillValue Empty { get; } = new XlsxDifferentialFillValue(null, null, null);

    internal XlsxDifferentialFillValue(XlsxPatternFill pattern)
        : this(pattern, null, null)
    {
    }

    internal XlsxDifferentialFillValue(XlsxFillFormatValue fill)
        : this(fill.Pattern, fill.LinearGradient, fill.PathGradient)
    {
    }

    private XlsxDifferentialFillValue(XlsxPatternFill? pattern, XlsxLinearGradientFill? linearGradient, XlsxPathGradientFill? pathGradient)
    {
        Pattern = pattern;
        LinearGradient = linearGradient;
        PathGradient = pathGradient;
    }

    internal XlsxPatternFill? Pattern { get; }

    internal XlsxLinearGradientFill? LinearGradient { get; }

    internal XlsxPathGradientFill? PathGradient { get; }
}
