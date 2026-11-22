using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.Hyperlinks;

public interface IXlsxHyperlinks : IEnumerable<XlsxHyperlink>
{
    /// <summary>
    /// Remove the hyperlink from a worksheet. Doesn't throw if hyperlinks is
    /// not attached to a worksheet.
    /// </summary>
    /// <remarks>
    /// If hyperlink range uses a <see cref="XlsxThemeColor.Hyperlink">hyperlink
    /// theme color</see>, the style is reset to the sheet style font color.
    /// The <see cref="IXlsxFontBase.Underline"/> is also set to sheet style
    /// underline.
    /// </remarks>
    /// <param name="hyperlink">Hyperlink to remove.</param>
    /// <returns><c>true</c> if hyperlink was part of the worksheet and was
    ///   removed. <c>false</c> otherwise.</returns>
    bool Delete(XlsxHyperlink hyperlink);

    /// <summary>
    /// Delete a hyperlink defined for a single cell. It doesn't delete
    /// hyperlinks that cover the cell.
    /// </summary>
    /// <remarks>
    /// If hyperlink range uses a <see cref="XlsxThemeColor.Hyperlink">hyperlink
    /// theme color</see>, the style is reset to the sheet style font color.
    /// The <see cref="IXlsxFontBase.Underline"/> is also set to sheet style
    /// underline.
    /// </remarks>
    /// <param name="address">Address of the cell.</param>
    /// <returns><c>true</c> if there was such hyperlink and was deleted.
    ///   <c>false</c> otherwise.</returns>
    bool Delete(IXlsxAddress address);

    /// <summary>
    /// Get a hyperlink for a single cell.
    /// </summary>
    /// <param name="address">Address of the cell.</param>
    /// <exception cref="KeyNotFoundException">Cell doesn't have a hyperlink.</exception>
    XlsxHyperlink Get(IXlsxAddress address);

    /// <summary>
    /// Get a hyperlink for a single cell.
    /// </summary>
    /// <param name="address">Address of the cell.</param>
    /// <param name="hyperlink">Found hyperlink.</param>
    /// <returns><c>true</c> if there was a hyperlink for <paramref name="address"/>, <c>false</c> otherwise.</returns>
    bool TryGet(IXlsxAddress address, [NotNullWhen(true)] out XlsxHyperlink? hyperlink);
}
