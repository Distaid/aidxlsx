using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// Source of data for a <see cref="XlsxPivotCache"/> that takes data from external workbook.
/// </summary>
internal sealed class XlsxPivotSourceExternalWorkbook : IXlsxPivotSource
{
    /// <summary>
    /// External workbook relId. If relationships of cache definition changes, make sure to either keep same or update it.
    /// </summary>
    internal string RelId { get; }

    /// <summary>
    /// Are source data in external workbook defined by a <see cref="TableOrName"/> or by <see cref="Area">cell area</see>.
    /// </summary>
    [MemberNotNullWhen(true, nameof(TableOrName))]
    [MemberNotNullWhen(false, nameof(Area))]
    internal bool UsesName => TableOrName is not null;

    /// <summary>
    /// A table or defined name in an external workbook that contains source data.
    /// </summary>
    internal string? TableOrName { get; }

    /// <summary>
    /// An area in an external workbook that contains source data.
    /// </summary>
    internal SheetArea? Area { get; }

    internal XlsxPivotSourceExternalWorkbook(string relId, SheetArea area)
    {
        RelId = relId;
        Area = area;
    }

    internal XlsxPivotSourceExternalWorkbook(string relId, string tableOrName)
    {
        RelId = relId;
        TableOrName = tableOrName;
    }

    public bool TryGetSource(XlsxWorkbook workbook, out XlsxWorksheet? sheet, out Area? sheetArea)
    {
        throw new NotImplementedException("External workbook source is not supported.");
    }

    public bool Equals(IXlsxPivotSource? otherSource)
    {
        var other = otherSource as XlsxPivotSourceExternalWorkbook;
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        // Two same RelIds could in theory point to different workbooks. I am not supporting
        // external sources for now anyway, so no unification through equality.
        return false;
    }

    public override bool Equals(object? other)
    {
        return other is IXlsxPivotSource source && Equals(source);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(RelId, Area).GetHashCode();
    }
}
