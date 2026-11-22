using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

/// <summary>
/// An API object to modify number format of a dxf.
/// </summary>
internal class XlsxDxfNumberFormat : IXlsxNumberFormat
{
    private readonly XlsxDxFormat _parent;

    internal XlsxDxfNumberFormat(XlsxDxFormat parent)
    {
        _parent = parent;
    }

    /// <inheritdoc />
    int IXlsxNumberFormatBase.NumberFormatId
    {
        get => NumberFormatId;
        set => NumberFormatId = value;
    }

    /// <inheritdoc />
    string IXlsxNumberFormatBase.Format
    {
        get => Format;
        set => Format = value;
    }

    private int NumberFormatId
    {
        get
        {
            var numberFormat = _parent.Resolve(static x => x.NumberFormat, static x => x);
            if (numberFormat is null)
                return XlsxPredefinedFormat.General;

            return XlsxPredefinedFormat.NumberFormatIds.GetValueOrDefault(new XlsxNumberFormat(numberFormat), -1);
        }
        set
        {
            if (!XlsxPredefinedFormat.FormatCodes.TryGetValue(value, out var format))
                throw new ArgumentOutOfRangeException($"Only predefined format is permitted. Use nested enums/members of {nameof(XlsxPredefinedFormat)}.");

            Format = format;
        }
    }

    private string Format
    {
        get
        {
            var format = _parent.Resolve(static x => x.NumberFormat, static x => x);
            if (format is null)
                return XlsxPredefinedFormat.FormatCodes[XlsxPredefinedFormat.General];

            return format;
        }
        set => _parent.ModifyNumberFormat(XlsxNumberFormat.Parse(value));
    }

    IXlsxStyle IXlsxNumberFormat.SetNumberFormatId(int value)
    {
        NumberFormatId = value;
        return _parent;
    }

    IXlsxStyle IXlsxNumberFormat.SetFormat(string value)
    {
        Format = value;
        return _parent;
    }

    bool IEquatable<IXlsxNumberFormatBase>.Equals(IXlsxNumberFormatBase? other)
    {
        if (other is null)
            return false;

        return Format == other.Format;
    }

    internal void SetValue(IXlsxNumberFormat value)
    {
        Format = value.Format;
    }
}
