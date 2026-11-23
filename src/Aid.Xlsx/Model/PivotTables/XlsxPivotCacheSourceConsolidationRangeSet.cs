using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// One of ranges that form a <see cref="XlsxPivotSourceConsolidation"/> source for a <see cref="XlsxPivotCache"/>.
/// </summary>
internal class XlsxPivotCacheSourceConsolidationRangeSet
{
    /// <summary>
    /// Indexes into the <see cref="XlsxPivotSourceConsolidation.Pages"/>. If the value is null
    /// and page filter exists, it is displayed as a blank. There can be at most 4 indexes, because
    /// there can be at most 4 page filters.
    /// </summary>
    public IReadOnlyList<uint?> Indexes { get; init; } = [];

    /// <summary>
    /// If range set is from another workbook, a relationship id to the workbook from cache definition.
    /// </summary>
    internal string? RelId { get; init; }

    [MemberNotNullWhen(true, nameof(TableOrName))]
    [MemberNotNullWhen(false, nameof(Area))]
    internal bool UsesName => TableOrName is not null;

    internal string? TableOrName { get; init; }

    internal SheetArea? Area { get; init; }
}
