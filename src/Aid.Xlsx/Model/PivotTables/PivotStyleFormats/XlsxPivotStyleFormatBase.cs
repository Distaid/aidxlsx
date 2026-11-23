using Aid.Xlsx.Model.Styles;

namespace Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

/// <summary>
/// A base class for pivot styling API. It has takes a selected <see cref="XlsxPivotArea"/>
/// and applies the style using <c>.Style*</c> API. The derived classes are responsible for
/// exposing API so user can define an area and then create the desired area (from what user
/// specified) through <see cref="GetCurrentArea"/> method.
/// </summary>
internal abstract class XlsxPivotStyleFormatBase(XlsxPivotTable pivotTable) : IXlsxPivotStyleFormat
{
    protected readonly XlsxPivotTable PivotTable = pivotTable;

    #region IXlsxPivotStyleFormat members

    public XlsxPivotStyleFormatElement AppliesTo { get; init; } = XlsxPivotStyleFormatElement.Data;

    public IXlsxStyle Style
    {
        get => Format;
        set => Format.SetStyle(value);
    }

    #endregion IXlsxPivotStyleFormat members

    // TODO Styles: Ensure that each pivot area is there only once in a pivot table. Ensure it on load and during modifications.
    internal XlsxDxFormat Format => new(PivotTable.Worksheet.Workbook.Styles, GetFormats().First());

    internal abstract XlsxPivotArea GetCurrentArea();

    internal abstract bool Filter(XlsxPivotArea area);

    private IEnumerable<XlsxPivotFormat> GetFormats()
    {
        var exists = false;
        foreach (var format in PivotTable.Formats)
        {
            if (format.Action == XlsxPivotFormatAction.Formatting && Filter(format.PivotArea))
            {
                exists = true;
                yield return format;
            }
        }

        if (!exists)
        {
            var format = new XlsxPivotFormat(GetCurrentArea())
            {
                FormatValue = null
            };
            PivotTable.AddFormat(format);
            yield return format;
        }
    }
}
