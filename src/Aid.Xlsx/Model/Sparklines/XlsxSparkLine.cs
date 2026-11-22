using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.Sparklines;

/// <summary>
/// An API object for a sparkline. It doesn't contain any data, only a link to the point of
/// the sparkline and does operations through <see cref="XlsxSparklineGroup"/>. It uses the
/// cell location as an anchor and if it is no longer valid, the group should throw.
/// </summary>
internal class XlsxSparkline : IXlsxSparkline
{
    private readonly XlsxSparklineGroup _sparklineGroup;
    private Point _location;

    internal XlsxSparkline(XlsxSparklineGroup sparklineGroup, Point location)
    {
        _sparklineGroup = sparklineGroup;
        _location = location;
    }

    public IXlsxCell Location
    {
        get => _sparklineGroup.GetLocation(_location);
        set => SetLocation(value);
    }

    public IXlsxRange? SourceData
    {
        get => _sparklineGroup.GetSparklineSourceData(_location);
        set => SetSourceData(value);
    }

    public IXlsxSparklineGroup SparklineGroup => _sparklineGroup;

    public IXlsxSparkline SetLocation(IXlsxCell newLocation)
    {
        if (newLocation is null)
            throw new ArgumentNullException(nameof(newLocation));

        if (newLocation.Worksheet != SparklineGroup.Worksheet)
            throw new ArgumentException("Cannot move the sparkline to a different worksheet");

        var destination = Point.FromCell(newLocation);
        _sparklineGroup.MoveSparkline(_location, destination);
        _location = destination;
        return this;
    }

    public IXlsxSparkline SetSourceData(IXlsxRange? sourceDataRange)
    {
        _sparklineGroup.SetSparklineSourceData(_location, sourceDataRange);
        return this;
    }
}
