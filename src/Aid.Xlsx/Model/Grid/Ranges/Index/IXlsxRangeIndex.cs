using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Model.Grid.Ranges.Index;

/// <summary>
/// Interface for the engine aimed to speed-up the search for the range intersections.
/// </summary>
internal interface IXlsxRangeIndex
{
    bool Add(IXlsxAddressable range);

    bool Remove(IXlsxRangeAddress rangeAddress);

    int RemoveAll(Predicate<IXlsxAddressable>? predicate = null);

    IEnumerable<IXlsxAddressable> GetIntersectedRanges(XlsxRangeAddress rangeAddress);

    IEnumerable<IXlsxAddressable> GetIntersectedRanges(XlsxAddress address);

    IEnumerable<IXlsxAddressable> GetAll();

    bool Intersects(in XlsxRangeAddress rangeAddress);

    bool Contains(in XlsxAddress address);

    bool MatchesType(XlsxRangeType rangeType);
}

internal interface IXlsxRangeIndex<T> : IXlsxRangeIndex
    where T : IXlsxAddressable
{
    bool Add(T range);

    int RemoveAll(Predicate<T>? predicate = null);

    new IEnumerable<T> GetIntersectedRanges(XlsxRangeAddress rangeAddress);

    new IEnumerable<T> GetIntersectedRanges(XlsxAddress address);

    new IEnumerable<T> GetAll();
}
