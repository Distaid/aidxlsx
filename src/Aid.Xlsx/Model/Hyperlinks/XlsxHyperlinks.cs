using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Hyperlinks;

internal class XlsxHyperlinks : IXlsxHyperlinks, ISheetListener
{
    private readonly XlsxWorksheet _worksheet;

    /// <summary>
    /// XlsxHyperlink doesn't contain range, it is user created and only then it is associated with an area in a sheet.
    /// </summary>
    private readonly List<(XlsxHyperlink Link, Area Area)> _hyperlinks = [];
    private readonly RTree<XlsxHyperlink> _areaIndex = new();
    private readonly Dictionary<XlsxHyperlink, Area> _linkIndex = [];

    private delegate (bool Success, Area? RepositionedArea) RepositionFunc(Area hyperlinkArea);

    internal XlsxHyperlinks(XlsxWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    internal string WorksheetName => _worksheet.Name;

    #region ISheetListener

    void ISheetListener.OnInsertAreaAndShiftDown(XlsxWorksheet sheet, Area insertedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryInsertAreaAndShiftDown(insertedArea, out Area? newHlArea);
            return (success, newHlArea);
        });
    }

    void ISheetListener.OnInsertAreaAndShiftRight(XlsxWorksheet sheet, Area insertedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryInsertAreaAndShiftRight(insertedArea, out Area? newHlArea);
            return (success, newHlArea);
        });
    }

    void ISheetListener.OnDeleteAreaAndShiftLeft(XlsxWorksheet sheet, Area deletedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryDeleteAreaAndShiftLeft(deletedArea, out Area? newHlArea);
            return (success, newHlArea);
        });
    }

    void ISheetListener.OnDeleteAreaAndShiftUp(XlsxWorksheet sheet, Area deletedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryDeleteAreaAndShiftUp(deletedArea, out Area? newHlArea);
            return (success, newHlArea);
        });
    }

    private void RepositionOnChange(XlsxWorksheet sheet, RepositionFunc reposition)
    {
        if (sheet != _worksheet)
            return;

        // Styles are responsibility of style slice => only shift areas
        foreach ((XlsxHyperlink? link, Area linkArea) in _hyperlinks.ToArray())
        {
            (bool success, Area? newLinkArea) = reposition(linkArea);
            if (!success)
                continue; // Partial cover, don't move.

            if (linkArea == newLinkArea)
                continue; // Nothing changed

            Remove(link);
            if (newLinkArea is not null)
                Add(newLinkArea.Value, link);
        }
    }

    #endregion ISheetListener

    public IEnumerator<XlsxHyperlink> GetEnumerator()
    {
        // Enumerate in same order it was loaded and will be saved
        return _hyperlinks.Select(static x => x.Link).GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <inheritdoc />
    public bool Delete(XlsxHyperlink hyperlink)
    {
        if (!Remove(hyperlink, out Area linkArea))
            return false;

        ClearHyperlinkStyle(linkArea);
        return true;
    }

    /// <inheritdoc />
    public bool Delete(IXlsxAddress address)
    {
        if (address.Worksheet is not null && address.Worksheet != _worksheet)
            return false;

        var cellPoint = Point.FromAddress(address);
        if (!TryGet(cellPoint, out XlsxHyperlink? cellLink))
            return false;

        Remove(cellLink);
        ClearHyperlinkStyle(cellPoint);
        return true;
    }

    /// <inheritdoc />
    public XlsxHyperlink Get(IXlsxAddress address)
    {
        if (address.Worksheet is not null && address.Worksheet != _worksheet)
            throw new KeyNotFoundException("Address is for a different sheet.");

        var point = Point.FromAddress(address);
        if (!TryGet(point, out XlsxHyperlink? link))
            throw new KeyNotFoundException($"No hyperlink is defined for cell {point}.");

        return link;
    }

    /// <inheritdoc />
    public bool TryGet(IXlsxAddress address, [NotNullWhen(true)] out XlsxHyperlink? hyperlink)
    {
        if (address.Worksheet is not null && address.Worksheet != _worksheet)
        {
            hyperlink = null;
            return false;
        }

        var point = Point.FromAddress(address);
        return TryGet(point, out hyperlink);
    }

    internal bool HasHyperlink(Point point)
    {
        var areaNodes = new List<RTree<XlsxHyperlink>.Node>();
        return _areaIndex.GetNodes(point, areaNodes).Count > 0;
    }

    /// <summary>
    /// Set a hyperlink of a single cell. Doesn't modify style, ignores hyperlinks with areas that
    /// cover the cell.
    /// </summary>
    internal void SetCellHyperlink(Point point, XlsxHyperlink? link)
    {
        // We only care about links defined for individual cell, not any link that covers the cell
        var pointNodes = new List<RTree<XlsxHyperlink>.Node>();
        _areaIndex.GetNodes(point, pointNodes);
        foreach (RTree<XlsxHyperlink>.Node existingLink in pointNodes)
            Remove(existingLink.Data);

        if (link is null)
            return;

        Add(point, link);
    }

    internal bool TryGet(Point point, [NotNullWhen(true)] out XlsxHyperlink? hyperlink)
    {
        var cellArea = new Area(point);
        var areaNodes = new List<RTree<XlsxHyperlink>.Node>();
        _areaIndex.GetNodes(cellArea, areaNodes);

        if (areaNodes.Count == 0)
        {
            hyperlink = null;
            return false;
        }

        if (areaNodes.Count == 1)
        {
            hyperlink = areaNodes[0].Data;
            return true;
        }

        // There are multiple areas for the point. When hyperlink areas overlap, Excel opens
        // the last one. So it is likely the correct one. But take a random one (areaNodes are
        // not guaranteed to be in correct order), because this API is just beyond any hope and
        // will be completely scrapped ASAP.
        hyperlink = areaNodes[^1].Data;
        return true;
    }

    internal XlsxCell? GetCell(XlsxHyperlink hyperlink)
    {
        if (!_linkIndex.TryGetValue(hyperlink, out Area area))
            return null;

        return new XlsxCell(_worksheet, area.FirstPoint);
    }

    private void Add(Area linkArea, XlsxHyperlink link)
    {
        if (link.Container is not null && link.Container != this)
            throw new InvalidOperationException("Hyperlink is attached to a different worksheet. Either remove it from the original worksheet or create a new hyperlink.");

        if (_linkIndex.ContainsKey(link))
            return;

        _linkIndex.Add(link, linkArea);
        _areaIndex.Insert(new RTree<XlsxHyperlink>.Node(linkArea, link));
        _hyperlinks.Add((link, linkArea));
        link.Container = this;
        Debug.Assert(_hyperlinks.Count == _linkIndex.Count);
        Debug.Assert(_hyperlinks.Count == _areaIndex.Count);
    }

    private void Remove(XlsxHyperlink link)
    {
        Remove(link, out _);
    }

    private bool Remove(XlsxHyperlink link, out Area area)
    {
        if (!_linkIndex.Remove(link, out area))
            return false;

        _areaIndex.Delete(new RTree<XlsxHyperlink>.Node(area, link));
        _hyperlinks.RemoveAll(x => x.Link == link);
        link.Container = null;
        Debug.Assert(_hyperlinks.Count == _linkIndex.Count);
        Debug.Assert(_hyperlinks.Count == _areaIndex.Count);
        return true;
    }

    private void ClearHyperlinkStyle(Area range)
    {
        XlsxFontFormatValue worksheetFont = _worksheet.GetFormat().Font;
        XlsxColor sheetColor = worksheetFont.Color;
        XlsxFontUnderlineValues sheetUnderline = worksheetFont.Underline;
        foreach (Point point in range)
        {
            XlsxCell? cell = _worksheet.GetCell(point);
            if (cell is null)
                continue;

            if (cell.Style.Font.FontColor.Equals(XlsxColor.FromTheme(XlsxThemeColor.Hyperlink)))
                cell.Style.Font.FontColor = sheetColor;

            cell.Style.Font.Underline = sheetUnderline;
        }
    }
}
