using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// API object for dxf hidden behind IXlsxStyle.IXlsxFill interface.
/// </summary>
internal class XlsxDxfFillFormat : IXlsxFill
{
    private static readonly XlsxPatternFill _default = new()
    {
        PatternType = XlsxFillPatternValues.None,
        BackgroundColor = XlsxColor.Automatic,
        PatternColor = XlsxColor.Automatic
    };

    private readonly XlsxDxFormat _parent;

    internal XlsxDxfFillFormat(XlsxDxFormat parent)
    {
        _parent = parent;
    }

    XlsxColor IXlsxFill.BackgroundColor
    {
        get => Resolve(static fill => fill.Pattern?.BackgroundColor, _default.BackgroundColor);
        set => Modify(static (fill, bgColor) =>
        {
            var currentPattern = fill.Pattern ?? _default;
            var newPattern = currentPattern.WithModifiedBgColor(bgColor);
            return new XlsxDifferentialFillValue(newPattern);
        }, value);
    }

    XlsxColor IXlsxFill.PatternColor
    {
        get => Resolve(static fill => fill.Pattern?.PatternColor, _default.PatternColor);
        set => Modify(static (fill, patternColor) =>
        {
            var pattern = fill.Pattern ?? _default;
            return new XlsxDifferentialFillValue(pattern with { PatternColor = patternColor });
        }, value);
    }

    XlsxFillPatternValues IXlsxFill.PatternType
    {
        get => Resolve(static fill => fill.Pattern?.PatternType, _default.PatternType);
        set => Modify(static (fill, patternType) =>
        {
            var pattern = fill.Pattern ?? _default;
            var newPattern = pattern.WithModifiedPattern(patternType);
            return new XlsxDifferentialFillValue(newPattern);
        }, value);
    }

    IXlsxStyle IXlsxFill.SetBackgroundColor(XlsxColor value)
    {
        (this as IXlsxFill).BackgroundColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxFill.SetPatternColor(XlsxColor value)
    {
        (this as IXlsxFill).PatternColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxFill.SetPatternType(XlsxFillPatternValues value)
    {
        (this as IXlsxFill).PatternType = value;
        return _parent;
    }

    bool IEquatable<IXlsxFill>.Equals(IXlsxFill? other)
    {
        throw new NotSupportedException();
    }

    internal void SetValue(IXlsxFill value)
    {
        // The original should be valid and consistent.
        Modify(static (_, patternFill) => new XlsxDifferentialFillValue(new XlsxPatternFill
        {
            PatternType = patternFill.PatternType,
            PatternColor = patternFill.PatternColor,
            BackgroundColor = patternFill.BackgroundColor,
        }), value);
    }

    private T Resolve<T>(Func<XlsxDifferentialFillValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Fill, getProperty) ?? defaultValue;
    }

    private T Resolve<T>(Func<XlsxDifferentialFillValue, T?> getProperty, T defaultValue)
        where T : class
    {
        return _parent.Resolve(static format => format.Fill, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XlsxDifferentialFillValue, T, XlsxDifferentialFillValue> modify, T value)
    {
        _parent.ModifyFill(modify, value);
    }
}
