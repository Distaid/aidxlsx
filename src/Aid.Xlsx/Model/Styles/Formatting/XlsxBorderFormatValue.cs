namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// A border format master record.
/// </summary>
/// <remarks>
/// Many XML attributes are used only for dxf and thus are not represented in this structure (vertical, horizontal, outline).
/// </remarks>
internal record XlsxBorderFormatValue
{
    internal static readonly XlsxBorderFormatValue None = new()
    {
        Left = XlsxBorderLine.None,
        Right = XlsxBorderLine.None,
        Top = XlsxBorderLine.None,
        Bottom = XlsxBorderLine.None,
        Diagonal = XlsxBorderLine.None,
        DiagonalUp = false,
        DiagonalDown = false,
    };

    public required XlsxBorderLine Left { get; init; }

    public required XlsxBorderLine Right { get; init; }

    public required XlsxBorderLine Top { get; init; }

    public required XlsxBorderLine Bottom { get; init; }

    /// <summary>
    /// Used when <see cref="DiagonalUp"/> or <see cref="DiagonalDown"/> are set. It's not possible
    /// to have different style for up/down diagonal.
    /// </summary>
    public required XlsxBorderLine Diagonal { get; init; }

    public required bool DiagonalUp { get; init; }

    public required bool DiagonalDown { get; init; }

    internal static XlsxBorderFormatValue FromDxf(XlsxDifferentialBorderValue dxfBorder)
    {
        return new XlsxBorderFormatValue
        {
            Left = dxfBorder.Left ?? XlsxBorderLine.None,
            Right = dxfBorder.Right ?? XlsxBorderLine.None,
            Top = dxfBorder.Top ?? XlsxBorderLine.None,
            Bottom = dxfBorder.Bottom ?? XlsxBorderLine.None,
            Diagonal = dxfBorder.Diagonal ?? XlsxBorderLine.None,
            DiagonalUp = dxfBorder.DiagonalUp,
            DiagonalDown = dxfBorder.DiagonalDown,
        };
    }
}
