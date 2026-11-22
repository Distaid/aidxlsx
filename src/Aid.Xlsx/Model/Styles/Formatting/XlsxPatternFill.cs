using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles.Formatting;

/// <summary>
/// A pattern fill format.
/// </summary>
internal record XlsxPatternFill
{
    /// <summary>
    /// Foreground color of the pattern (i.e., if pattern is a grid, this is the color to draw
    /// lines of the grid).
    /// </summary>
    internal required XlsxColor PatternColor { get; init; }

    /// <summary>
    /// Background color of the pattern.
    /// </summary>
    internal required XlsxColor BackgroundColor { get; init; }

    /// <summary>
    /// Shape of the pattern.
    /// </summary>
    internal required XlsxFillPatternValues PatternType { get; init; }

    internal XlsxPatternFill WithModifiedPattern(XlsxFillPatternValues patternType)
    {
        if (PatternType == XlsxFillPatternValues.None && patternType != XlsxFillPatternValues.None)
        {
            // Keep the original behavior. Just having a non-none pattern type doesn't actually
            // affects color of cells visually, it needs at least background color.
            return this with
            {
                BackgroundColor = XlsxColor.FromTheme(XlsxThemeColor.Text1),
                PatternType = patternType
            };
        }

        return this with { PatternType = patternType };
    }

    internal XlsxPatternFill WithModifiedBgColor(XlsxColor newBgColor)
    {
        XlsxFillPatternValues newPattern;
        if (newBgColor.IsAuto && PatternType == XlsxFillPatternValues.Solid)
        {
            // Fill is single color, and that color will be is transparent -> there is no fill.
            newPattern = XlsxFillPatternValues.None;
        }
        else if (!newBgColor.IsAuto && PatternType == XlsxFillPatternValues.None)
        {
            // Color is not transparent, but fill is none -> make it solid
            newPattern = XlsxFillPatternValues.Solid;
        }
        else
        {
            newPattern = PatternType;
        }

        return new XlsxPatternFill
        {
            PatternType = newPattern,
            PatternColor = PatternColor,
            BackgroundColor = newBgColor
        };
    }
}
