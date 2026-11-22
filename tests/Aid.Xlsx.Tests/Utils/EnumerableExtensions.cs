using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Tests;

internal static class EnumerableExtensions
{
    public static string ToSpaceList(this IEnumerable<IXlsxRange> ranges, bool includeSheet = false)
    {
        return string.Join(" ", ranges.Select(r => r.RangeAddress.ToString(XlsxReferenceStyle.A1, includeSheet)));
    }

    public static string ToSpaceList(this IEnumerable<Area> areas)
    {
        return string.Join(" ", areas);
    }
}
