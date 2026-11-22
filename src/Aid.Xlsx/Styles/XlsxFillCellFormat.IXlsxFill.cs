using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Styles;

internal sealed partial class XlsxFillCellFormat : IXlsxFill
{
    XlsxColor IXlsxFill.BackgroundColor
    {
        get => BackgroundColor;
        set => BackgroundColor = value;
    }

    XlsxColor IXlsxFill.PatternColor
    {
        get => PatternColor;
        set => PatternColor = value;
    }

    XlsxFillPatternValues IXlsxFill.PatternType
    {
        get => PatternType;
        set => PatternType = value;
    }

    IXlsxStyle IXlsxFill.SetBackgroundColor(XlsxColor value)
    {
        BackgroundColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxFill.SetPatternColor(XlsxColor value)
    {
        PatternColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxFill.SetPatternType(XlsxFillPatternValues value)
    {
        PatternType = value;
        return _parent;
    }

    bool IEquatable<IXlsxFill>.Equals(IXlsxFill? other)
    {
        if (other is null)
            return false;

        // This is a "business" equality, i.e. will both fills look the same.
        // This is gradient fill, other can only represent pattern, regardless of what it actually is.
        var isPatternFill = _parent.Resolve(static x => x.Fill.Pattern) is not null;
        if (!isPatternFill)
            return false;

        if (!HasFill(this) && !HasFill(other))
            return true;

        if (PatternType != other.PatternType)
            return false;

        if (BackgroundColor != other.BackgroundColor)
            return false;

        var patternColor = UsesPatternColor(this) ? PatternColor : null;
        var otherPatternColor = UsesPatternColor(other) ? other.PatternColor : null;
        if (patternColor != otherPatternColor)
            return false;

        return true;

        static bool HasFill(IXlsxFill fill)
        {
            var patternType = fill.PatternType;
            if (patternType == XlsxFillPatternValues.None)
                return false;

            if (patternType == XlsxFillPatternValues.Solid && fill.BackgroundColor.IsAuto)
                return false;

            return true;
        }

        static bool UsesPatternColor(IXlsxFill fill)
        {
            return fill.PatternType is not XlsxFillPatternValues.None and not XlsxFillPatternValues.Solid;
        }
    }
}
