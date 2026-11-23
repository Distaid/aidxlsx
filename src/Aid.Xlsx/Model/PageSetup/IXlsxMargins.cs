#nullable disable

namespace Aid.Xlsx.Model.PageSetup;

public interface IXlsxMargins
{
    /// <summary>Gets or sets the Left margin.</summary>
    /// <value>The Left margin.</value>
    double Left { get; set; }

    /// <summary>Gets or sets the Right margin.</summary>
    /// <value>The Right margin.</value>
    double Right { get; set; }

    /// <summary>Gets or sets the Top margin.</summary>
    /// <value>The Top margin.</value>
    double Top { get; set; }

    /// <summary>Gets or sets the Bottom margin.</summary>
    /// <value>The Bottom margin.</value>
    double Bottom { get; set; }

    /// <summary>Gets or sets the Header margin.</summary>
    /// <value>The Header margin.</value>
    double Header { get; set; }

    /// <summary>Gets or sets the Footer margin.</summary>
    /// <value>The Footer margin.</value>
    double Footer { get; set; }

    IXlsxMargins SetLeft(double value);
    IXlsxMargins SetRight(double value);
    IXlsxMargins SetTop(double value);
    IXlsxMargins SetBottom(double value);
    IXlsxMargins SetHeader(double value);
    IXlsxMargins SetFooter(double value);

}
