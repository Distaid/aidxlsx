#nullable disable

using Aid.Xlsx.Model.RichText;

namespace Aid.Xlsx.Model.PageSetup;

public enum XlsxHFPredefinedText
{
    PageNumber, NumberOfPages, Date, Time, FullPath, Path, File, SheetName
}

/// <summary>
/// An enum specifying pages where the content of the <see cref="IXlsxHFItem"/> be displayed.
/// </summary>
public enum XlsxHFOccurrence
{
    /// <summary>
    /// Header or footer on the all pages.
    /// </summary>
    AllPages,

    /// <summary>
    /// Header or footer on the odd pages. It is necessary to enable <see cref="IXlsxPageSetup.DifferentOddEvenPagesOnHF"/>
    /// to display the different heading/footer on odd pages. If there is a different header/footer on the first
    /// page and the <see cref="IXlsxPageSetup.DifferentFirstPageOnHF"/> is enabled, the first page will not use the odd
    /// page setting.
    /// </summary>
    OddPages,

    /// <summary>
    /// Header or footer on the even pages. It is necessary to enable <see cref="IXlsxPageSetup.DifferentOddEvenPagesOnHF"/>
    /// to display the different heading/footer on even pages.
    /// </summary>
    EvenPages,

    /// <summary>
    /// Header or footer on the first page. It is necessary to enable <see cref="IXlsxPageSetup.DifferentFirstPageOnHF"/>
    /// to display the different heading/footer on the first page.
    /// </summary>
    FirstPage
}

public interface IXlsxHFItem : IXlsxWithRichString
{
    /// <summary>
    /// Gets the text of the specified header/footer occurrence.
    /// </summary>
    /// <param name="occurrence">The occurrence.</param>
    String GetText(XlsxHFOccurrence occurrence);

    /// <summary>
    /// Adds the given predefined text to this header/footer item.
    /// </summary>
    /// <param name="predefinedText">The predefined text to add to this header/footer item.</param>
    IXlsxRichString AddText(XlsxHFPredefinedText predefinedText);

    /// <summary>
    /// Adds the given text to this header/footer item.
    /// </summary>
    /// <param name="text">The text to add to this header/footer item.</param>
    /// <param name="occurrence">The occurrence for the text.</param>
    IXlsxRichString AddText(String text, XlsxHFOccurrence occurrence);

    /// <summary>
    /// Adds the given predefined text to this header/footer item.
    /// </summary>
    /// <param name="predefinedText">The predefined text to add to this header/footer item.</param>
    /// <param name="occurrence">The occurrence for the predefined text.</param>
    IXlsxRichString AddText(XlsxHFPredefinedText predefinedText, XlsxHFOccurrence occurrence);

    /// <summary>Clears the text/formats of this header/footer item.</summary>
    /// <param name="occurrence">The occurrence to clear.</param>
    void Clear(XlsxHFOccurrence occurrence = XlsxHFOccurrence.AllPages);

    IXlsxRichString AddImage(String imagePath, XlsxHFOccurrence occurrence = XlsxHFOccurrence.AllPages);
}
