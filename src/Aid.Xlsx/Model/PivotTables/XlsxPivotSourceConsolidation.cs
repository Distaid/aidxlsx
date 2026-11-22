using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// Source of data for a <see cref="XlsxPivotCache"/> that takes uses a union of multiple scenarios in the workbook to
/// create data.
/// </summary>
internal sealed class XlsxPivotSourceConsolidation : IXlsxPivotSource
{
    /// <summary>
    /// Will application automatically create additional page filter in addition to the <see cref="Pages"/>.
    /// </summary>
    internal bool AutoPage { get; init; }

    /// <summary>
    /// <para>
    /// Custom page filters that toggle whether to display data from a particular
    /// <see cref="XlsxPivotCacheSourceConsolidationRangeSet">range set</see>.
    /// There can be 0..4 page filters. Each can have a different combination
    /// of range sets.
    /// </para>
    /// <para>
    /// Example:
    /// <example>
    /// The range sets are months and one page is <em>Q1</em>,<em>Q2</em>,<em>Q3</em>,<em>Q4</em>
    /// and second page filter is <em>Last month of quarter</em> and <em>Other months</em>. These
    /// page items are referenced by <see cref="XlsxPivotCacheSourceConsolidationRangeSet.Indexes"/>.
    /// </example>
    /// </para>
    /// </summary>
    public IReadOnlyList<XlsxPivotCacheSourceConsolidationPage> Pages { get; init; } = Array.Empty<XlsxPivotCacheSourceConsolidationPage>();

    /// <summary>
    /// Range sets that consists the cache source.
    /// </summary>
    public IReadOnlyList<XlsxPivotCacheSourceConsolidationRangeSet> RangeSets { get; init; } = Array.Empty<XlsxPivotCacheSourceConsolidationRangeSet>();

    public bool Equals(IXlsxPivotSource? otherSource)
    {
        var other = otherSource as XlsxPivotSourceConsolidation;
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        // This source should likely never be unified, so when there are two instances, mark them as different.
        return false;
    }

    public bool TryGetSource(XlsxWorkbook workbook, out XlsxWorksheet? sheet, out Area? sheetArea)
    {
        throw new NotImplementedException("Consolidation pivot cache data source is not supported.");
    }
}
