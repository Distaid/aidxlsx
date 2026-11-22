using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// An API object to modify protection of a dxf.
/// </summary>
internal class XlsxDxfProtectionFormat : IXlsxProtection
{
    private readonly XlsxDxFormat _parent;

    public XlsxDxfProtectionFormat(XlsxDxFormat parent)
    {
        _parent = parent;
    }

    bool IXlsxProtection.Locked
    {
        get => Locked;
        set => Locked = value;
    }

    bool IXlsxProtection.Hidden
    {
        get => Hidden;
        set => Hidden = value;
    }

    private bool Locked
    {
        get => Resolve(p => p.Locked, true);
        set => Modify(static (protection, locked) => protection with { Locked = locked }, value);
    }

    private bool Hidden
    {
        get => Resolve(p => p.Hidden, false);
        set => Modify(static (protection, hidden) => protection with { Hidden = hidden }, value);
    }

    IXlsxStyle IXlsxProtection.SetLocked()
    {
        Locked = true;
        return _parent;
    }

    IXlsxStyle IXlsxProtection.SetLocked(bool value)
    {
        Locked = value;
        return _parent;
    }

    IXlsxStyle IXlsxProtection.SetHidden()
    {
        Hidden = true;
        return _parent;
    }

    IXlsxStyle IXlsxProtection.SetHidden(bool value)
    {
        Hidden = value;
        return _parent;
    }

    bool IEquatable<IXlsxProtection>.Equals(IXlsxProtection? other)
    {
        if (other is null)
            return false;

        return Locked == other.Locked && Hidden == other.Hidden;
    }

    internal void SetValue(IXlsxProtection value)
    {
        Modify((_, other) => new XlsxDifferentialProtectionValue
        {
            Locked = other.Locked,
            Hidden = other.Hidden
        }, value);
    }

    private T Resolve<T>(Func<XlsxDifferentialProtectionValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Protection, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XlsxDifferentialProtectionValue, T, XlsxDifferentialProtectionValue> modify, T value)
    {
        _parent.ModifyProtection(modify, value);
    }
}
