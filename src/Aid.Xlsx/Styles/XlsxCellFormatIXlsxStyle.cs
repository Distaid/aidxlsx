namespace Aid.Xlsx.Styles;

/// <summary>
/// Methods and properties of <see cref="IXlsxStyle"/>. The <see cref="XlsxCellFormat"/> has many
/// properties with same name, but different type, so the interface is explicitly implemented.
/// </summary>
internal partial class XlsxCellFormat : IXlsxStyle
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
        set => Font.SetFont(value);
    }

    bool IXlsxStyle.IncludeQuotePrefix
    {
        get => IncludeQuotePrefix;
        set => IncludeQuotePrefix = value;
    }

    IXlsxNumberFormat IXlsxStyle.NumberFormat
    {
        get => NumberFormat;
        set => NumberFormat.SetNumberFormat(value.Format);
    }

    IXlsxProtection IXlsxStyle.Protection
    {
        get => Protection;
        set => Protection.SetValue(value);
    }

    IXlsxStyle IXlsxStyle.SetIncludeQuotePrefix(bool includeQuotePrefix)
    {
        IncludeQuotePrefix = includeQuotePrefix;
        return this;
    }

    bool IEquatable<IXlsxStyle>.Equals(IXlsxStyle? other)
    {
        if (other is null)
            return false;

        // The API object for each component implement IEquitable<IXL..>
        if (!NumberFormat.Equals(other.NumberFormat))
            return false;

        if (!Font.Equals(other.Font))
            return false;

        if (!IncludeQuotePrefix.Equals(other.IncludeQuotePrefix))
            return false;

        if (!Fill.Equals(other.Fill))
            return false;

        if (!Border.Equals(other.Border))
            return false;

        if (!Alignment.Equals(other.Alignment))
            return false;

        if (!Protection.Equals(other.Protection))
            return false;

        return true;
    }

    /// <summary>
    /// A helper method that is used when a style if copied from one object to another.
    /// For example, <c>rangeApi.Style = someOtherApi.Style</c>.
    /// </summary>
    internal void SetStyle(IXlsxStyle value)
    {
        if (value is not XlsxCellFormat cellFormat)
            throw new NotSupportedException("Can only copy cell format style.");

        var otherCellFormat = cellFormat._formatValue.Resolve();
        var registeredCellFormat = _workbook.Styles.GetRegisteredCellFormat(otherCellFormat);
        ModifyFormat((_, cellForamt) => cellForamt, registeredCellFormat);
    }
}
