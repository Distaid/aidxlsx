using Aid.Xlsx.Model.ConditionalFormats;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// Specification of conditional formatting of a pivot table.
/// </summary>
internal class XlsxPivotConditionalFormat
{
    private readonly List<XlsxPivotArea> _area = new();

    internal XlsxPivotConditionalFormat(XlsxConditionalFormat format)
    {
        Format = format;
    }

    /// <summary>
    /// An option to display in GUI on how to update <see cref="Areas"/>.
    /// </summary>
    internal XlsxPivotCfScope Scope { get; init; } = XlsxPivotCfScope.SelectedCells;

    /// <summary>
    /// A rule that determines how should CF be applied to <see cref="Areas"/>.
    /// </summary>
    /// <remarks>Doesn't seem to work, Excel has no dialogue, nothing found on web and Excel tries
    ///     to repair on row/column values. Avoid if possible.</remarks>
    internal XlsxPivotCfRuleType Type { get; init; } = XlsxPivotCfRuleType.None;

    /// <summary>
    /// Areas of pivot table the rule should be applied. The areas are projected to the sheet
    /// <see cref="XlsxConditionalFormat.Ranges"/> that Excel actually uses to display CF.
    /// </summary>
    internal IReadOnlyList<XlsxPivotArea> Areas => _area;

    /// <summary>
    /// Conditional format applied to the <see cref="Areas"/>.
    /// </summary>
    /// <remarks>
    /// The <see cref="XlsxConditionalFormat.Priority"/> of the format is used as a identifier used
    /// to connect pivot CF element and sheet CF element. Pivot CF is ultimately part of sheet CFs
    /// and the priority determines order of CF application (note that CF has
    /// <see cref="XlsxConditionalFormat.StopIfTrue"/> flag).
    /// </remarks>
    internal XlsxConditionalFormat Format { get; }

    internal void AddArea(XlsxPivotArea pivotArea)
    {
        _area.Add(pivotArea);
    }
}
