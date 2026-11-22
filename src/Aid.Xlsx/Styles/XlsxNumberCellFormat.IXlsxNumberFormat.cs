using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

internal sealed partial class XlsxNumberCellFormat : IXlsxNumberFormat
{
    int IXlsxNumberFormatBase.NumberFormatId
    {
        get => NumberFormatId;
        set => NumberFormatId = value;
    }

    string IXlsxNumberFormatBase.Format
    {
        get => Format;
        set => Format = XlsxNumberFormat.Parse(value);
    }

    bool IEquatable<IXlsxNumberFormatBase>.Equals(IXlsxNumberFormatBase? other)
    {
        if (other is null)
            return false;

        return other.Format == Format;
    }

    IXlsxStyle IXlsxNumberFormat.SetNumberFormatId(int value)
    {
        NumberFormatId = value;
        return _parent;
    }

    IXlsxStyle IXlsxNumberFormat.SetFormat(string value)
    {
        Format = XlsxNumberFormat.Parse(value);
        return _parent;
    }
}
