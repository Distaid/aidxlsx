using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles.Formatting;

/// <summary>
/// A linear gradient fill format of a cell.
/// </summary>
internal record XlsxLinearGradientFill
{
    /// <summary>
    /// The key is a position of color gradient from position 0 to 1. If stops are not available
    /// for values for 0 and 1, the gradient uses black color for them. The collection can be empty.
    /// </summary>
    public required IReadOnlyDictionary<FractionOfOne, XlsxColor> Stops { get; init; }

    public required double Degrees { get; init; }
}
