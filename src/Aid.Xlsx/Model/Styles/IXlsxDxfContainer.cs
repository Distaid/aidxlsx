using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// An interface used by object that use differential formatting, e.g. conditional formatting or
/// pivot tables.
/// </summary>
internal interface IXlsxDxfContainer
{
    /// <summary>
    /// The differential format value.
    /// </summary>
    /// <remarks>
    /// The value is optional, because attribute in XML is optional.
    /// </remarks>
    XlsxDxfValue? FormatValue { get; set; }
}
