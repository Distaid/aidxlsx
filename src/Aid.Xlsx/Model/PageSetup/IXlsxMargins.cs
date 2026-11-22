#nullable disable

namespace Aid.Xlsx.Model.PageSetup;

public interface IXlsxMargins
{
    /// <summary>Gets or sets the Left margin.</summary>
    /// <value>The Left margin.</value>
    Double Left { get; set; }

    /// <summary>Gets or sets the Right margin.</summary>
    /// <value>The Right margin.</value>
    Double Right { get; set; }

    /// <summary>Gets or sets the Top margin.</summary>
    /// <value>The Top margin.</value>
    Double Top { get; set; }

    /// <summary>Gets or sets the Bottom margin.</summary>
    /// <value>The Bottom margin.</value>
    Double Bottom { get; set; }

    /// <summary>Gets or sets the Header margin.</summary>
    /// <value>The Header margin.</value>
    Double Header { get; set; }

    /// <summary>Gets or sets the Footer margin.</summary>
    /// <value>The Footer margin.</value>
    Double Footer { get; set; }

    IXlsxMargins SetLeft(Double value);
    IXlsxMargins SetRight(Double value);
    IXlsxMargins SetTop(Double value);
    IXlsxMargins SetBottom(Double value);
    IXlsxMargins SetHeader(Double value);
    IXlsxMargins SetFooter(Double value);

}
