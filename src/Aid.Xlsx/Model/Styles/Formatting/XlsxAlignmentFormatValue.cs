namespace Aid.Xlsx.Model.Styles.Formatting;

internal record XlsxAlignmentFormatValue
{
    private readonly int _indent;

    /// <summary>
    /// Default values of alignment properties in XML part. If a value is missing in XML this one is used instead.
    /// </summary>
    internal static XlsxAlignmentFormatValue Default { get; } = new()
    {
        Horizontal = XlsxAlignmentHorizontalValues.General,
        Vertical = XlsxAlignmentVerticalValues.Bottom,
        TextRotation = TextRotation.None,
        WrapText = false,
        Indent = 0,
        RelativeIndent = 0,
        JustifyLastLine = false,
        ShrinkToFit = false,
        ReadingOrder = XlsxAlignmentReadingOrderValues.ContextDependent,
    };

    public required XlsxAlignmentHorizontalValues Horizontal { get; init; }

    public required XlsxAlignmentVerticalValues Vertical { get; init; }

    public required TextRotation TextRotation { get; init; }

    /// <summary>
    /// Should the text be line-wrapped within the cell?
    /// </summary>
    public required bool WrapText { get; init; }

    /// <summary>
    /// Indicates number of 3*spaces (of the normal style font) of indentation for text in a cell.
    /// </summary>
    public required int Indent
    {
        get => _indent;
        init
        {
            if (value is < 0 or > 255)
                throw new ArgumentOutOfRangeException(nameof(value), value, "Indent must be between 0 and 255.");

            _indent = value;
        }
    }

    /// <summary>
    /// Relative indentation for dxf. Indicates number of spaces to indent text in a cell.
    /// </summary>
    public required int RelativeIndent { get; init; }

    public required bool JustifyLastLine { get; init; }

    public required bool ShrinkToFit { get; init; }

    public required XlsxAlignmentReadingOrderValues ReadingOrder { get; init; }
}
