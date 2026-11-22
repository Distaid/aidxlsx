using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles.Formatting;

/// <summary>
/// Border line properties.
/// </summary>
internal readonly record struct XlsxBorderLine(XlsxColor Color, XlsxBorderStyleValues Style)
{
    /// <summary>
    /// Default XML attribute values.
    /// </summary>
    internal static readonly XlsxBorderLine None = new(XlsxColor.Automatic, XlsxBorderStyleValues.None);

    /// <summary>
    /// Is the border line visible?
    /// </summary>
    internal bool IsVisible => Style != XlsxBorderStyleValues.None;
}
