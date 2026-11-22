using Aid.Xlsx.Styles.Formatting;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.PivotTables;

/// <summary>
/// A description of formatting that should be applied to a <see cref="XlsxPivotTable"/>.
/// </summary>
internal class XlsxPivotFormat : IXlsxDxfContainer
{
    internal XlsxPivotFormat(XlsxPivotArea pivotArea)
    {
        PivotArea = pivotArea;
    }

    /// <summary>
    /// Pivot area that should be formatted.
    /// </summary>
    internal XlsxPivotArea PivotArea { get; }

    /// <summary>
    /// Should the formatting (determined by <see cref="FormatValue"/>) be applied or not?
    /// </summary>
    internal XlsxPivotFormatAction Action { get; init; } = XlsxPivotFormatAction.Formatting;

    /// <summary>
    /// Differential formatting to apply to the <see cref="PivotArea"/>. It can be empty, e.g. if
    /// <see cref="Action"/> is blank.
    /// </summary>
    public XlsxDxfValue? FormatValue { get; set; }
}
