using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

internal sealed partial class XlsxProtectionCellFormat
{
    private readonly XlsxCellFormat _parent;

    internal XlsxProtectionCellFormat(XlsxCellFormat parent)
    {
        _parent = parent;
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxProtection other && (this as IEquatable<IXlsxProtection>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetValue(IXlsxProtection value)
    {
        Modify(static (_, other) => new XlsxProtectionFormatValue
        {
            Hidden = other.Hidden,
            Locked = other.Locked
        }, value);
    }

    private T Resolve<T>(Func<XlsxCellFormatValue, T> selector)
    {
        return _parent.Resolve(selector);
    }

    private void Modify<TProperty>(Func<XlsxProtectionFormatValue, TProperty, XlsxProtectionFormatValue> modifyProtection, TProperty value)
    {
        _parent.ModifyProtection(modifyProtection, value);
    }
}
