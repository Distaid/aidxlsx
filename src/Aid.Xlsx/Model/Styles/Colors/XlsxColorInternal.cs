namespace Aid.Xlsx.Model.Styles.Colors;

public partial class XlsxColor
{
    internal XlsxColorKey Key { get; }

    private XlsxColor() : this(new XlsxColorKey() { ColorType = XlsxColorType.Automatic })
    {
        HasValue = false;
    }

    internal XlsxColor(XlsxColorKey key)
    {
        Key = key;
        HasValue = true;
    }

    /// <summary>
    /// Lower case color type for exception messages.
    /// </summary>
    private string LcColorType => ColorType.ToString().ToLowerInvariant();
}
