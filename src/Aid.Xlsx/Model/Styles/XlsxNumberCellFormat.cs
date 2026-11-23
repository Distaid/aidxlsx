using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// An API object to modify number format of a <see cref="XlsxCellFormat">cell format</see>.
/// </summary>
internal sealed partial class XlsxNumberCellFormat
{
    private readonly XlsxCellFormat _parent;

    internal XlsxNumberCellFormat(XlsxCellFormat parent)
    {
        _parent = parent;
    }

    private int NumberFormatId
    {
        get
        {
            var numberFormat = _parent.Resolve(static x => x.NumberFormat);
            return XlsxPredefinedFormat.NumberFormatIds.GetValueOrDefault(numberFormat, -1);
        }
        set
        {
            if (!XlsxPredefinedFormat.FormatCodes.TryGetValue(value, out var format))
                throw new ArgumentOutOfRangeException(nameof(value), value, $"Only predefined format is permitted. Use nested enums/members of {nameof(XlsxPredefinedFormat)}.");

            Format = format;
        }
    }

    private XlsxNumberFormat Format
    {
        get => _parent.Resolve(static x => x.NumberFormat);
        set => _parent.ModifyNumberFormat(value);
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxNumberFormatBase other && (this as IEquatable<IXlsxNumberFormatBase>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetNumberFormat(string numberFormat)
    {
        Format = XlsxNumberFormat.Parse(numberFormat);
    }
}
