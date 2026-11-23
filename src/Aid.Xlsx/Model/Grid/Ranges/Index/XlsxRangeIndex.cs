using Aid.Xlsx.Infrastructure.Patterns;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges.Index;

/// <summary>
/// Implementation of <see cref="IXlsxRangeIndex"/> internally using QuadTree.
/// </summary>
internal abstract class XlsxRangeIndex : IXlsxRangeIndex
{
    #region Public Constructors

    public XlsxRangeIndex(IXlsxWorksheet worksheet)
    {
        _worksheet = worksheet;
        _rangeList = [];
        ((XlsxWorksheet)_worksheet).RegisterRangeIndex(this);
    }

    #endregion Public Constructors

    #region Public Methods

    public abstract bool MatchesType(XlsxRangeType rangeType);

    public bool Add(IXlsxAddressable range)
    {
        ArgumentNullException.ThrowIfNull(range);

        if (!range.RangeAddress.IsValid)
            throw new ArgumentException("Range is invalid");

        CheckWorksheet(range.RangeAddress.Worksheet);

        _count++;
        if (_count < MinimumCountForIndexing)
        {
            if (_rangeList.Any(r => r == range))
                return false;

            _rangeList.Add(range);
            return true;
        }

        if (_quadTree == null)
            InitializeTree();

        return _quadTree!.Add(range);
    }

    public bool Contains(in XlsxAddress address)
    {
        CheckWorksheet(address.Worksheet);

        if (_quadTree == null)
        {
            var addr = address;
            return _rangeList.Any(r => r.RangeAddress.Contains(addr));
        }

        return _quadTree.GetIntersectedRanges(address).Any();
    }

    public IEnumerable<IXlsxAddressable> GetAll()
    {
        if (_quadTree == null)
        {
            return _rangeList.AsEnumerable();
        }

        return _quadTree.GetAll();
    }

    public IEnumerable<IXlsxAddressable> GetIntersectedRanges(XlsxRangeAddress rangeAddress)
    {
        CheckWorksheet(rangeAddress.Worksheet);

        if (_quadTree == null)
        {
            return _rangeList.Where(r => r.RangeAddress.Intersects(rangeAddress));
        }

        return _quadTree.GetIntersectedRanges(rangeAddress);
    }

    public IEnumerable<IXlsxAddressable> GetIntersectedRanges(XlsxAddress address)
    {
        CheckWorksheet(address.Worksheet);

        if (_quadTree == null)
        {
            return _rangeList.Where(r => r.RangeAddress.Contains(address));
        }

        return _quadTree.GetIntersectedRanges(address);
    }

    public bool Intersects(in XlsxRangeAddress rangeAddress)
    {
        CheckWorksheet(rangeAddress.Worksheet);

        if (_quadTree == null)
        {
            var addr = rangeAddress;
            return _rangeList.Any(r => r.RangeAddress.Intersects(addr));
        }

        return _quadTree.GetIntersectedRanges(rangeAddress).Any();
    }

    public bool Remove(IXlsxRangeAddress rangeAddress)
    {
        ArgumentNullException.ThrowIfNull(rangeAddress);

        CheckWorksheet(rangeAddress.Worksheet);

        if (_quadTree == null)
        {
            return _rangeList.RemoveAll(r => Equals(r.RangeAddress, rangeAddress)) > 0;
        }

        return _quadTree.Remove(rangeAddress);
    }

    public int RemoveAll(Predicate<IXlsxAddressable>? predicate = null)
    {
        predicate ??= (_ => true);

        if (_quadTree == null)
        {
            return _rangeList.RemoveAll(predicate);
        }

        return _quadTree.RemoveAll(predicate).Count();
    }

    #endregion Public Methods

    #region Private Fields

    /// <summary>
    /// The minimum number of ranges to be included into a QuadTree. Until it is reached the ranges
    /// are added into a simple list to minimize the overhead of searching intersections on small collections.
    /// </summary>
    private const int MinimumCountForIndexing = 20;

    /// <summary>
    /// A collection of ranges used before the QuadTree is initialized (until <see cref="MinimumCountForIndexing"/>
    /// is reached.
    /// </summary>
    protected readonly List<IXlsxAddressable> _rangeList;

    private readonly IXlsxWorksheet _worksheet;
    private int _count = 0;
    protected Quadrant? _quadTree;

    #endregion Private Fields

    #region Private Methods

    private void CheckWorksheet(IXlsxWorksheet? worksheet)
    {
        if (worksheet != _worksheet)
            throw new ArgumentException("Range belongs to a different worksheet");
    }

    private void InitializeTree()
    {
        _quadTree = CreateQuadTree();
        _rangeList.ForEach(r => _quadTree.Add(r));
        _rangeList.Clear();
    }

    protected virtual Quadrant CreateQuadTree()
    {
        return new Quadrant();
    }

    #endregion Private Methods
}

/// <summary>
/// Generic version of <see cref="XlsxRangeIndex"/>.
/// </summary>
internal class XlsxRangeIndex<T> : XlsxRangeIndex, IXlsxRangeIndex<T>
    where T : IXlsxAddressable
{
    public XlsxRangeIndex(IXlsxWorksheet worksheet) : base(worksheet)
    {
    }

    public bool Add(T range)
    {
        return base.Add(range);
    }

    public int RemoveAll(Predicate<T>? predicate)
    {
        predicate ??= (_ => true);

        return base.RemoveAll(r => predicate((T)r));
    }

    public new IEnumerable<T> GetIntersectedRanges(XlsxRangeAddress rangeAddress)
    {
        return base.GetIntersectedRanges(rangeAddress).Cast<T>();
    }

    public new IEnumerable<T> GetIntersectedRanges(XlsxAddress address)
    {
        return base.GetIntersectedRanges(address).Cast<T>();
    }

    public override bool MatchesType(XlsxRangeType rangeType)
    {
        var innerType = typeof(T);

        if (innerType == typeof(IXlsxRangeBase) ||
            innerType == typeof(XlsxRangeBase))
            return true;

        return rangeType switch
        {
            XlsxRangeType.Range => innerType == typeof(IXlsxRange) ||
                                   innerType == typeof(XlsxRange),
            XlsxRangeType.Column => innerType == typeof(IXlsxColumn) ||
                                   innerType == typeof(XlsxColumn),
            XlsxRangeType.Row => innerType == typeof(IXlsxRow) ||
                                   innerType == typeof(XlsxRow),
            XlsxRangeType.RangeColumn => innerType == typeof(IXlsxRangeColumn) ||
                                   innerType == typeof(XlsxRangeColumn),
            XlsxRangeType.RangeRow => innerType == typeof(IXlsxRangeRow) ||
                                   innerType == typeof(XlsxRangeRow),
            XlsxRangeType.Table => innerType == typeof(IXlsxTable) ||
                                   innerType == typeof(XlsxTable),
            XlsxRangeType.Worksheet => innerType == typeof(IXlsxWorksheet) ||
                                   innerType == typeof(XlsxWorksheet),
            _ => throw new NotImplementedException(nameof(rangeType)),
        };
    }

    public new IEnumerable<T> GetAll()
    {
        return base.GetAll().Cast<T>();
    }

    protected override Quadrant CreateQuadTree()
    {
        return new Quadrant<T>();
    }
}
