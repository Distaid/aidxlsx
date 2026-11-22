using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

internal partial class XlsxProtectionCellFormat : IXlsxProtection
{
    bool IXlsxProtection.Locked
    {
        get => Resolve(static style => style.Protection.Locked);
        set => Modify(static (protection, locked) => protection with { Locked = locked }, value);
    }

    bool IXlsxProtection.Hidden
    {
        get => Resolve(static style => style.Protection.Hidden);
        set => Modify(static (protection, hidden) => protection with { Hidden = hidden }, value);
    }

    IXlsxStyle IXlsxProtection.SetLocked()
    {
        return (this as IXlsxProtection).SetLocked(true);
    }

    IXlsxStyle IXlsxProtection.SetLocked(bool value)
    {
        (this as IXlsxProtection).Locked = value;
        return _parent;
    }

    IXlsxStyle IXlsxProtection.SetHidden()
    {
        return (this as IXlsxProtection).SetHidden(true);
    }

    IXlsxStyle IXlsxProtection.SetHidden(bool value)
    {
        (this as IXlsxProtection).Hidden = value;
        return _parent;
    }

    bool IEquatable<IXlsxProtection>.Equals(IXlsxProtection? other)
    {
        if (other is null)
            return false;

        var protection = Resolve(static style => style.Protection);
        return protection.Locked == other.Locked && protection.Hidden == other.Hidden;
    }
}
