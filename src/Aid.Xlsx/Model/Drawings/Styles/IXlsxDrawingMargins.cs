namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingMargins
{
    bool Automatic { get; set; }

    /// <summary>
    /// Left margin in inches.
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// Right margin in inches.
    /// </summary>
    double Right { get; set; }

    /// <summary>
    /// Top margin in inches.
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// Bottom margin in inches.
    /// </summary>
    double Bottom { get; set; }

    /// <summary>
    /// Set <see cref="Left"/>, <see cref="Top"/>, <see cref="Right"/>, <see cref="Bottom"/> margins at once.
    /// </summary>
    double All { set; }

    IXlsxDrawingStyle SetAutomatic(); IXlsxDrawingStyle SetAutomatic(bool value);
    IXlsxDrawingStyle SetLeft(double value);
    IXlsxDrawingStyle SetRight(double value);
    IXlsxDrawingStyle SetTop(double value);
    IXlsxDrawingStyle SetBottom(double value);
    IXlsxDrawingStyle SetAll(double value);

}
