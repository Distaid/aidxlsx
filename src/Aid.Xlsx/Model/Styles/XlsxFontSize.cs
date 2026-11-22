namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// Size of a font stored in twips. Storing font size as double causes various problems with
/// equality of font size. Per MS-OI29500: Office converts the points provided to twips, and
/// rounding may occur when writing sz@val back to SpreadsheetML files.
/// </summary>
internal readonly record struct XlsxFontSize : IEquatable<double>
{
    public XlsxFontSize(short twips)
    {
        if (twips is < 20 or > 8191)
            throw new ArgumentOutOfRangeException(nameof(twips), "Font size must be between 1 and 409.55 points.");

        Twips = twips;
    }

    public static XlsxFontSize FromPoints(double sizeInPoints)
    {
        var twips = Math.Round(sizeInPoints * 20, MidpointRounding.AwayFromZero);
        return new XlsxFontSize(checked((short)twips));
    }

    public short Twips { get; }

    /// <summary>
    /// Font size converted to points. Can have rounding issues, so use only when necessary.
    /// </summary>
    public double Points => Twips / 20.0;

    public bool Equals(double other)
    {
        return other.Equals(Points);
    }
}
