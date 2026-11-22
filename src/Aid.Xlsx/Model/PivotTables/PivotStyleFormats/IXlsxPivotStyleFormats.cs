// Keep this file CodeMaid organised and cleaned

namespace Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

/// <summary>
/// An API for setting style of parts consisting of <see cref="XlsxPivotStyleFormatElement"/>, e.g. grand
/// totals. The enumerator enumerates only existing formats, it doesn't add them.
/// </summary>
public interface IXlsxPivotStyleFormats : IEnumerable<IXlsxPivotStyleFormat>
{
    /// <summary>
    /// Get styling object for specified <paramref name="element"/>.
    /// </summary>
    /// <param name="element">Which part do we want style for?</param>
    /// <returns>An API to inspect/modify style of the <paramref name="element"/>.</returns>
    /// <exception cref="ArgumentException">When <see cref="XlsxPivotStyleFormatElement.None"/> is
    ///     passed as an argument.</exception>
    IXlsxPivotStyleFormat ForElement(XlsxPivotStyleFormatElement element);
}
