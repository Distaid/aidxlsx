namespace Aid.Xlsx.Drawings.Styles
{
    public interface IXlsxDrawingMargins
    {
        Boolean Automatic { get; set; }

        /// <summary>
        /// Left margin in inches.
        /// </summary>
        Double Left { get; set; }

        /// <summary>
        /// Right margin in inches.
        /// </summary>
        Double Right { get; set; }

        /// <summary>
        /// Top margin in inches.
        /// </summary>
        Double Top { get; set; }

        /// <summary>
        /// Bottom margin in inches.
        /// </summary>
        Double Bottom { get; set; }

        /// <summary>
        /// Set <see cref="Left"/>, <see cref="Top"/>, <see cref="Right"/>, <see cref="Bottom"/> margins at once.
        /// </summary>
        Double All { set; }

        IXlsxDrawingStyle SetAutomatic(); IXlsxDrawingStyle SetAutomatic(Boolean value);
        IXlsxDrawingStyle SetLeft(Double value);
        IXlsxDrawingStyle SetRight(Double value);
        IXlsxDrawingStyle SetTop(Double value);
        IXlsxDrawingStyle SetBottom(Double value);
        IXlsxDrawingStyle SetAll(Double value);

    }
}
