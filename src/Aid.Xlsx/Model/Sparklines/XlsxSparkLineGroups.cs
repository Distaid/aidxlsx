// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Sparklines;

internal class XlsxSparklineGroups : IXlsxSparklineGroups, IEnumerable<XlsxSparklineGroup>
{
    private readonly XlsxWorksheet _worksheet;
    private readonly List<XlsxSparklineGroup> _sparklineGroups = [];

    public XlsxSparklineGroups(XlsxWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    public IXlsxWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Add empty sparkline group.
    /// </summary>
    internal XlsxSparklineGroup Add()
    {
        var emptyGroup = new XlsxSparklineGroup(Worksheet);
        Add(emptyGroup);
        return emptyGroup;
    }

    /// <summary>
    /// Add the sparkline group to the collection.
    /// </summary>
    /// <param name="sparklineGroup">The sparkline group to add to the collection</param>
    /// <returns>The same sparkline group</returns>
    public IXlsxSparklineGroup Add(IXlsxSparklineGroup sparklineGroup)
    {
        if (sparklineGroup.Worksheet != Worksheet)
            throw new ArgumentException("The specified sparkline group belongs to the different worksheet");

        _sparklineGroups.Add((XlsxSparklineGroup)sparklineGroup);
        return sparklineGroup;
    }

    public IXlsxSparklineGroup Add(string locationAddress, string sourceDataAddress)
    {
        ArgumentNullException.ThrowIfNull(locationAddress);

        ArgumentNullException.ThrowIfNull(sourceDataAddress);

        return Add(new XlsxSparklineGroup(Worksheet, locationAddress, sourceDataAddress));
    }

    public IXlsxSparklineGroup Add(IXlsxCell location, IXlsxRange sourceData)
    {
        ArgumentNullException.ThrowIfNull(location);

        return Add(new XlsxSparklineGroup(location, sourceData));
    }

    public IXlsxSparklineGroup Add(IXlsxRange locationRange, IXlsxRange sourceDataRange)
    {
        ArgumentNullException.ThrowIfNull(locationRange);

        ArgumentNullException.ThrowIfNull(sourceDataRange);

        return Add(new XlsxSparklineGroup(locationRange, sourceDataRange));
    }

    /// <summary>
    /// Add a copy of an existing sparkline group to the specified worksheet
    /// </summary>
    /// <param name="sparklineGroupToCopy">The sparkline group to copy</param>
    /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
    /// <returns>The new sparkline group added</returns>
    public IXlsxSparklineGroup AddCopy(IXlsxSparklineGroup sparklineGroupToCopy, IXlsxWorksheet targetWorksheet)
    {
        var sparklineGroup = new XlsxSparklineGroup(targetWorksheet, sparklineGroupToCopy);
        _sparklineGroups.Add(sparklineGroup);
        return sparklineGroup;
    }

    /// <summary>
    /// Copy this sparkline group to a different worksheet
    /// </summary>
    /// <param name="targetSheet">The worksheet to copy the sparkline group to</param>
    public void CopyTo(IXlsxWorksheet targetSheet)
    {
        foreach (var slg in this)
        {
            slg.CopyTo((XlsxWorksheet)targetSheet);
        }
    }

    /// <summary>
    /// Search for the first sparkline that is in the specified cell
    /// </summary>
    /// <param name="cell">The cell to find the sparkline for</param>
    /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
    public IXlsxSparkline? GetSparkline(IXlsxCell cell)
    {
        if (cell.Worksheet != _worksheet)
            return null;

        var location = Point.FromCell(cell);
        foreach (var sparklineGroup in _sparklineGroups)
        {
            if (sparklineGroup.TryGetSparkline(location, out var sparkline))
                return sparkline;
        }

        return null;
    }

    /// <summary>
    /// Find all sparklines located in a given range
    /// </summary>
    /// <param name="searchRange">The range to search</param>
    /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
    public IEnumerable<IXlsxSparkline> GetSparklines(IXlsxRangeBase searchRange)
    {
        return _sparklineGroups
            .SelectMany(g => g.GetSparklines(searchRange));
    }

    public IEnumerator<XlsxSparklineGroup> GetEnumerator()
    {
        return _sparklineGroups.GetEnumerator();
    }

    IEnumerator<IXlsxSparklineGroup> IEnumerable<IXlsxSparklineGroup>.GetEnumerator()
    {
        return GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// Remove all sparklines in the specified cell
    /// </summary>
    /// <param name="cell">The cell to remove sparklines from</param>
    public void Remove(IXlsxCell cell)
    {
        _sparklineGroups
            .AsParallel()
            .ForEach(g => g.Remove(cell));
    }

    public void Remove(IXlsxRangeBase range)
    {
        var sparklinesToRemove = _sparklineGroups
            .SelectMany(g => g)
            .Where(sparkline => range.Contains(sparkline.Location))
            .ToList();

        sparklinesToRemove.ForEach(Remove);
    }

    /// <summary>
    /// Remove the sparkline group from the worksheet
    /// </summary>
    /// <param name="sparklineGroup">The sparkline group to remove</param>
    public void Remove(IXlsxSparklineGroup sparklineGroup)
    {
        _sparklineGroups.Remove((XlsxSparklineGroup)sparklineGroup);
    }

    internal void Remove(Point location)
    {
        _sparklineGroups.ForEach(g => g.Remove(location));
    }

    /// <summary>
    /// Remove the sparkline from the worksheet
    /// </summary>
    /// <param name="sparkline">The sparkline to remove</param>
    private void Remove(IXlsxSparkline sparkline)
    {
        sparkline.SparklineGroup.Remove(sparkline);
    }

    /// <summary>
    /// Remove all sparkline groups and their contents from the worksheet.
    /// </summary>
    public void RemoveAll()
    {
        _sparklineGroups.Clear();
    }
}
