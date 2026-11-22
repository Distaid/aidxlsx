// Keep this file CodeMaid organised and cleaned
using System.Collections;

namespace Aid.Xlsx.PivotTables.PivotStyleFormats;

/// <summary>
/// An API for grand totals from <see cref="XlsxPivotTableStyleFormats"/>.
/// </summary>
internal class XlsxPivotStyleFormats : IXlsxPivotStyleFormats
{
    private readonly XlsxPivotTable _pivotTable;
    private readonly bool _isRowGrand;

    internal XlsxPivotStyleFormats(XlsxPivotTable pivotTable, bool isRowGrand)
    {
        _pivotTable = pivotTable;
        _isRowGrand = isRowGrand;
    }

    #region IXlsxPivotStyleFormats members

    public IXlsxPivotStyleFormat ForElement(XlsxPivotStyleFormatElement element)
    {
        if (element == XlsxPivotStyleFormatElement.None)
            throw new ArgumentException("Choose an enum value that represents an element", nameof(element));

        return GetPivotStyleFormatFor(element);
    }

    public IEnumerator<IXlsxPivotStyleFormat> GetEnumerator()
    {
        var elements = new[]
        {
            XlsxPivotStyleFormatElement.Label,
            XlsxPivotStyleFormatElement.Data,
            XlsxPivotStyleFormatElement.All,
        };

        foreach (var element in elements)
        {
            foreach (var format in _pivotTable.Formats)
            {
                if (AreaBelongsToGrandTotal(format.PivotArea, element))
                {
                    // Each pivot style format modifies all formats, so return only once per element.
                    yield return GetPivotStyleFormatFor(element);
                    break;
                }
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    #endregion IXlsxPivotStyleFormats members

    private XlsxPivotStyleFormat GetPivotStyleFormatFor(XlsxPivotStyleFormatElement element)
    {
        return new XlsxPivotStyleFormat(_pivotTable, FilterElement, ElementFactory)
        {
            AppliesTo = element
        };

        bool FilterElement(XlsxPivotArea pivotArea) => AreaBelongsToGrandTotal(pivotArea, element);
        XlsxPivotArea ElementFactory() => CreateGrandArea(element);
    }

    private bool AreaBelongsToGrandTotal(XlsxPivotArea area, XlsxPivotStyleFormatElement element)
    {
        return
            area.References.Count == 0 &&
            area.Field is null &&
            area.Type == XlsxPivotAreaType.Normal &&
            area.DataOnly == (element == XlsxPivotStyleFormatElement.Data) &&
            area.LabelOnly == (element == XlsxPivotStyleFormatElement.Label) &&
            area.GrandRow == _isRowGrand &&
            area.GrandCol == !_isRowGrand &&
            area.CacheIndex == false &&
            area.Offset is null &&
            !area.CollapsedLevelsAreSubtotals &&
            area.Axis is null &&
            area.FieldPosition is null;
    }

    private XlsxPivotArea CreateGrandArea(XlsxPivotStyleFormatElement element)
    {
        return new XlsxPivotArea
        {
            DataOnly = (element == XlsxPivotStyleFormatElement.Data),
            LabelOnly = (element == XlsxPivotStyleFormatElement.Label),
            GrandRow = _isRowGrand,
            GrandCol = !_isRowGrand,
        };
    }
}
