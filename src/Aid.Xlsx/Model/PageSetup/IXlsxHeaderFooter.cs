#nullable disable

namespace Aid.Xlsx.Model.PageSetup;

public enum XlsxHFMode { OddPagesOnly, OddAndEvenPages, Odd }
public interface IXlsxHeaderFooter
{
    /// <summary>
    /// Gets the left header/footer item.
    /// </summary>
    IXlsxHFItem Left { get; }

    /// <summary>
    /// Gets the middle header/footer item.
    /// </summary>
    IXlsxHFItem Center { get; }

    /// <summary>
    /// Gets the right header/footer item.
    /// </summary>
    IXlsxHFItem Right { get; }

    /// <summary>
    /// Gets the text of the specified header/footer occurrence.
    /// </summary>
    /// <param name="occurrence">The occurrence.</param>
    String GetText(XlsxHFOccurrence occurrence);

    IXlsxHeaderFooter Clear(XlsxHFOccurrence occurrence = XlsxHFOccurrence.AllPages);
}
