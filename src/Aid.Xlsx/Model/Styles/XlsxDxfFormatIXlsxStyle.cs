namespace Aid.Xlsx.Model.Styles;

internal partial class XlsxDxFormat : IXlsxStyle
{
    IXlsxAlignment IXlsxStyle.Alignment
    {
        get => Alignment;
        set => Alignment.SetValue(value);
    }

    IXlsxBorder IXlsxStyle.Border
    {
        get => Border;
        set => Border.SetValue(value);
    }

    IXlsxNumberFormat IXlsxStyle.DateFormat => NumberFormat;

    IXlsxFill IXlsxStyle.Fill
    {
        get => Fill;
        set => Fill.SetValue(value);
    }

    IXlsxFont IXlsxStyle.Font
    {
        get => Font;
        set => Font.SetValue(value);
    }

    bool IXlsxStyle.IncludeQuotePrefix
    {
        get => false;
        set => throw new NotSupportedException($"Differential format doesn't support {nameof(IXlsxStyle.IncludeQuotePrefix)}.");
    }

    IXlsxNumberFormat IXlsxStyle.NumberFormat
    {
        get => NumberFormat;
        set => NumberFormat.SetValue(value);
    }

    IXlsxProtection IXlsxStyle.Protection
    {
        get => Protection;
        set => Protection.SetValue(value);
    }

    IXlsxStyle IXlsxStyle.SetIncludeQuotePrefix(bool includeQuotePrefix)
    {
        (this as IXlsxStyle).IncludeQuotePrefix = includeQuotePrefix;
        return this;
    }

    bool IEquatable<IXlsxStyle>.Equals(IXlsxStyle? other)
    {
        throw new NotSupportedException();
    }
}
