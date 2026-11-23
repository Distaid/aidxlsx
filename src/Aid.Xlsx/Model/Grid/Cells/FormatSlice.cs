using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Grid.Cells;

internal class FormatSlice : ISlice
{
    private readonly Slice<XlsxCellFormatValue?> _slice = new();

    public bool IsEmpty => _slice.IsEmpty;

    public int MaxColumn => _slice.MaxColumn;

    public int MaxRow => _slice.MaxRow;

    public Dictionary<int, int>.KeyCollection UsedColumns => _slice.UsedColumns;

    public IEnumerable<int> UsedRows => _slice.UsedRows;

    public void Clear(Area area)
    {
        _slice.Clear(area);
    }

    public void DeleteAreaAndShiftLeft(Area areaToDelete)
    {
        _slice.DeleteAreaAndShiftLeft(areaToDelete);
    }

    public void DeleteAreaAndShiftUp(Area areaToDelete)
    {
        _slice.DeleteAreaAndShiftUp(areaToDelete);
    }

    public IEnumerator<Point> GetEnumerator(Area area, bool reverse = false)
    {
        return _slice.GetEnumerator(area, reverse);
    }

    public void InsertAreaAndShiftDown(Area areaToInsert)
    {
        _slice.InsertAreaAndShiftDown(areaToInsert);
    }

    public void InsertAreaAndShiftRight(Area areaToInsert)
    {
        _slice.InsertAreaAndShiftRight(areaToInsert);
    }

    public bool IsUsed(Point address)
    {
        return _slice.IsUsed(address);
    }

    public void Swap(Point sp1, Point sp2)
    {
        _slice.Swap(sp1, sp2);
    }

    public void Set(Point point, XlsxCellFormatValue? value)
    {
        _slice.Set(point, value);
    }

    internal void SetAll(Area area, XlsxCellFormatValue? value)
    {
        _slice.SetAll(area, value);
    }

    internal XlsxCellFormatValue? GetFormat(Point point)
    {
        return _slice[point];
    }

    internal void AddUsedFormat(HashSet<XlsxCellFormatValue> usedCellFormats)
    {
        IEnumerator<Point> enumerator = GetEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            if (_slice[enumerator.Current] is { } format)
                usedCellFormats.Add(format);
        }
    }
}
