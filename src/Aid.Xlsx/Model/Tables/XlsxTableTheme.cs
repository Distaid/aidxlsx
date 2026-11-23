using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Tables;

public sealed partial class XlsxTableTheme
{
    public static readonly XlsxTableTheme None = new("None");
    public static readonly XlsxTableTheme TableStyleMedium28 = new("TableStyleMedium28");
    public static readonly XlsxTableTheme TableStyleMedium27 = new("TableStyleMedium27");
    public static readonly XlsxTableTheme TableStyleMedium26 = new("TableStyleMedium26");
    public static readonly XlsxTableTheme TableStyleMedium25 = new("TableStyleMedium25");
    public static readonly XlsxTableTheme TableStyleMedium24 = new("TableStyleMedium24");
    public static readonly XlsxTableTheme TableStyleMedium23 = new("TableStyleMedium23");
    public static readonly XlsxTableTheme TableStyleMedium22 = new("TableStyleMedium22");
    public static readonly XlsxTableTheme TableStyleMedium21 = new("TableStyleMedium21");
    public static readonly XlsxTableTheme TableStyleMedium20 = new("TableStyleMedium20");
    public static readonly XlsxTableTheme TableStyleMedium19 = new("TableStyleMedium19");
    public static readonly XlsxTableTheme TableStyleMedium18 = new("TableStyleMedium18");
    public static readonly XlsxTableTheme TableStyleMedium17 = new("TableStyleMedium17");
    public static readonly XlsxTableTheme TableStyleMedium16 = new("TableStyleMedium16");
    public static readonly XlsxTableTheme TableStyleMedium15 = new("TableStyleMedium15");
    public static readonly XlsxTableTheme TableStyleMedium14 = new("TableStyleMedium14");
    public static readonly XlsxTableTheme TableStyleMedium13 = new("TableStyleMedium13");
    public static readonly XlsxTableTheme TableStyleMedium12 = new("TableStyleMedium12");
    public static readonly XlsxTableTheme TableStyleMedium11 = new("TableStyleMedium11");
    public static readonly XlsxTableTheme TableStyleMedium10 = new("TableStyleMedium10");
    public static readonly XlsxTableTheme TableStyleMedium9 = new("TableStyleMedium9");
    public static readonly XlsxTableTheme TableStyleMedium8 = new("TableStyleMedium8");
    public static readonly XlsxTableTheme TableStyleMedium7 = new("TableStyleMedium7");
    public static readonly XlsxTableTheme TableStyleMedium6 = new("TableStyleMedium6");
    public static readonly XlsxTableTheme TableStyleMedium5 = new("TableStyleMedium5");
    public static readonly XlsxTableTheme TableStyleMedium4 = new("TableStyleMedium4");
    public static readonly XlsxTableTheme TableStyleMedium3 = new("TableStyleMedium3");
    public static readonly XlsxTableTheme TableStyleMedium2 = new("TableStyleMedium2");
    public static readonly XlsxTableTheme TableStyleMedium1 = new("TableStyleMedium1");
    public static readonly XlsxTableTheme TableStyleLight21 = new("TableStyleLight21");
    public static readonly XlsxTableTheme TableStyleLight20 = new("TableStyleLight20");
    public static readonly XlsxTableTheme TableStyleLight19 = new("TableStyleLight19");
    public static readonly XlsxTableTheme TableStyleLight18 = new("TableStyleLight18");
    public static readonly XlsxTableTheme TableStyleLight17 = new("TableStyleLight17");
    public static readonly XlsxTableTheme TableStyleLight16 = new("TableStyleLight16");
    public static readonly XlsxTableTheme TableStyleLight15 = new("TableStyleLight15");
    public static readonly XlsxTableTheme TableStyleLight14 = new("TableStyleLight14");
    public static readonly XlsxTableTheme TableStyleLight13 = new("TableStyleLight13");
    public static readonly XlsxTableTheme TableStyleLight12 = new("TableStyleLight12");
    public static readonly XlsxTableTheme TableStyleLight11 = new("TableStyleLight11");
    public static readonly XlsxTableTheme TableStyleLight10 = new("TableStyleLight10");
    public static readonly XlsxTableTheme TableStyleLight9 = new("TableStyleLight9");
    public static readonly XlsxTableTheme TableStyleLight8 = new("TableStyleLight8");
    public static readonly XlsxTableTheme TableStyleLight7 = new("TableStyleLight7");
    public static readonly XlsxTableTheme TableStyleLight6 = new("TableStyleLight6");
    public static readonly XlsxTableTheme TableStyleLight5 = new("TableStyleLight5");
    public static readonly XlsxTableTheme TableStyleLight4 = new("TableStyleLight4");
    public static readonly XlsxTableTheme TableStyleLight3 = new("TableStyleLight3");
    public static readonly XlsxTableTheme TableStyleLight2 = new("TableStyleLight2");
    public static readonly XlsxTableTheme TableStyleLight1 = new("TableStyleLight1");
    public static readonly XlsxTableTheme TableStyleDark11 = new("TableStyleDark11");
    public static readonly XlsxTableTheme TableStyleDark10 = new("TableStyleDark10");
    public static readonly XlsxTableTheme TableStyleDark9 = new("TableStyleDark9");
    public static readonly XlsxTableTheme TableStyleDark8 = new("TableStyleDark8");
    public static readonly XlsxTableTheme TableStyleDark7 = new("TableStyleDark7");
    public static readonly XlsxTableTheme TableStyleDark6 = new("TableStyleDark6");
    public static readonly XlsxTableTheme TableStyleDark5 = new("TableStyleDark5");
    public static readonly XlsxTableTheme TableStyleDark4 = new("TableStyleDark4");
    public static readonly XlsxTableTheme TableStyleDark3 = new("TableStyleDark3");
    public static readonly XlsxTableTheme TableStyleDark2 = new("TableStyleDark2");
    public static readonly XlsxTableTheme TableStyleDark1 = new("TableStyleDark1");

    public string Name { get; }

    private readonly Dictionary<XlsxTableStyleRegionValues, XlsxDxfValue> _regionFormats = [];

    public XlsxTableTheme(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Value is empty.", nameof(name));

        this.Name = name;
    }

    public static IEnumerable<XlsxTableTheme> GetAllThemes() => BuiltInThemes.All;

    public static XlsxTableTheme? FromName(string name)
    {
        return name is not null && BuiltInThemes.ByName.TryGetValue(name, out XlsxTableTheme? theme) ? theme : null;
    }

    internal IReadOnlyDictionary<XlsxTableStyleRegionValues, XlsxDxfValue> RegionFormats => _regionFormats;

    /// <summary>
    /// A band size (i.e. how many rows does a stripe have) for a <see cref="XlsxTableStyleRegionValues.FirstRowStripe"/>.
    /// </summary>
    internal int RowStripe1BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many rows does a stripe have) for a <see cref="XlsxTableStyleRegionValues.SecondRowStripe"/>.
    /// </summary>
    internal int RowStripe2BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many column does a stripe have) for a <see cref="XlsxTableStyleRegionValues.FirstColumnStripe"/>.
    /// </summary>
    internal int ColumnStripe1BandSize { get; private set; } = 1;

    /// <summary>
    /// A band size (i.e. how many column does a stripe have) for a <see cref="XlsxTableStyleRegionValues.SecondColumnStripe"/>.
    /// </summary>
    internal int ColumnStripe2BandSize { get; private set; } = 1;

    #region Overrides

    public override bool Equals(object? obj)
    {
        var theme = obj as XlsxTableTheme;
        return theme is not null && Name.Equals(theme.Name);
    }

    public override int GetHashCode()
    {
        return this.Name.GetHashCode();
    }

    public override string ToString()
    {
        return this.Name;
    }

    #endregion Overrides

    internal void SetRegionFormat(XlsxTableStyleRegionValues region, XlsxDxfValue dxf, int bandSize = 1)
    {
        // Keep values excel compatible.
        if (bandSize is < 0 or > 9)
            throw new ArgumentOutOfRangeException(nameof(bandSize));

        // Band size has only meaning for stripe regions
        switch (region)
        {
            case XlsxTableStyleRegionValues.FirstRowStripe:
                RowStripe1BandSize = bandSize;
                break;
            case XlsxTableStyleRegionValues.SecondRowStripe:
                RowStripe2BandSize = bandSize;
                break;
            case XlsxTableStyleRegionValues.FirstColumnStripe:
                ColumnStripe1BandSize = bandSize;
                break;
            case XlsxTableStyleRegionValues.SecondColumnStripe:
                ColumnStripe2BandSize = bandSize;
                break;
        }

        _regionFormats[region] = dxf;
    }
}
