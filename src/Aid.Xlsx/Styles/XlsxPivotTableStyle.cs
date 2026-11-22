using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

/// <summary>
/// A named style for a pivot table.
/// </summary>
internal class XlsxPivotTableStyle
{
    private readonly Dictionary<XlsxPivotStyleRegionValues, XlsxDxfValue> _regionFormats = new();

    public XlsxPivotTableStyle(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Name can't be empty.");

        Name = name;
    }

    internal string Name { get; }

    internal IReadOnlyDictionary<XlsxPivotStyleRegionValues, XlsxDxfValue> RegionFormats => _regionFormats;

    /// <summary>
    /// A band size (i.e. how many rows does a stripe have) for a <see cref="XlsxPivotStyleRegionValues.FirstRowStripe"/>.
    /// </summary>
    internal int RowStripe1BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many rows does a stripe have) for a <see cref="XlsxPivotStyleRegionValues.SecondRowStripe"/>.
    /// </summary>
    internal int RowStripe2BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many column does a stripe have) for a <see cref="XlsxPivotStyleRegionValues.FirstColumnStripe"/>.
    /// </summary>
    internal int ColumnStripe1BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many column does a stripe have) for a <see cref="XlsxPivotStyleRegionValues.SecondColumnStripe"/>.
    /// </summary>
    internal int ColumnStripe2BandSize { get; private set; } = 1;

    internal void SetRegionFormat(XlsxPivotStyleRegionValues region, XlsxDxfValue dxf, int bandSize = 1)
    {
        // Keep values excel compatible.
        if (bandSize is < 0 or > 9)
            throw new ArgumentOutOfRangeException(nameof(bandSize));

        // Band size has only meaning for stripe regions
        switch (region)
        {
            case XlsxPivotStyleRegionValues.FirstRowStripe:
                RowStripe1BandSize = bandSize;
                break;
            case XlsxPivotStyleRegionValues.SecondRowStripe:
                RowStripe2BandSize = bandSize;
                break;
            case XlsxPivotStyleRegionValues.FirstColumnStripe:
                ColumnStripe1BandSize = bandSize;
                break;
            case XlsxPivotStyleRegionValues.SecondColumnStripe:
                ColumnStripe2BandSize = bandSize;
                break;
        }

        _regionFormats[region] = dxf;
    }
}
