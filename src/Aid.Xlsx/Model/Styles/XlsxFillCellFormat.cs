using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// An API object to modify fill of a cell range.
/// </summary>
/// <remarks>
/// Most of the various logic that adjusts background color or pattern based on changes to pattern
/// type or background color have been inherited from original implementation.
/// </remarks>
internal sealed partial class XlsxFillCellFormat
{
    private static readonly XlsxPatternFill PatternNone = new()
    {
        PatternType = XlsxFillPatternValues.None,
        PatternColor = XlsxColor.Automatic,
        BackgroundColor = XlsxColor.Automatic
    };

    private readonly XlsxCellFormat _parent;

    public XlsxFillCellFormat(XlsxCellFormat parent)
    {
        _parent = parent;
    }

    private XlsxColor BackgroundColor
    {
        get => _parent.Resolve(static x => (x.Fill.Pattern ?? PatternNone).BackgroundColor);
        set => _parent.ModifyFill(static (fill, bgColor) =>
        {
            var currentPattern = fill.Pattern ?? PatternNone;
            var newPattern = currentPattern.WithModifiedBgColor(bgColor);
            return new XlsxFillFormatValue(newPattern);
        }, value);
    }

    private XlsxColor PatternColor
    {
        get => _parent.Resolve(x => (x.Fill.Pattern ?? PatternNone).PatternColor);
        set => _parent.ModifyFill((fill, patternColor) =>
        {
            var pattern = fill.Pattern ?? PatternNone;
            return new XlsxFillFormatValue(pattern with { PatternColor = patternColor });
        }, value);
    }

    private XlsxFillPatternValues PatternType
    {
        get => _parent.Resolve(static x => (x.Fill.Pattern ?? PatternNone).PatternType);
        set => _parent.ModifyFill(static (fill, patternType) =>
        {
            var pattern = fill.Pattern ?? PatternNone;
            var newPattern = pattern.WithModifiedPattern(patternType);
            return new XlsxFillFormatValue(newPattern);
        }, value);
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxFill other && (this as IEquatable<IXlsxFill>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetValue(IXlsxFill value)
    {
        // No need for shenanigans with changing pattern fill or colors, the original should be valid and consistent.
        _parent.ModifyFill(static (_, patternFill) => new XlsxFillFormatValue(new XlsxPatternFill
        {
            PatternType = patternFill.PatternType,
            PatternColor = patternFill.PatternColor,
            BackgroundColor = patternFill.BackgroundColor,
        }), value);
    }
}
