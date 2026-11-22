namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// A border of differential format.
/// </summary>
/// <remarks>
/// Properties that are pivot/table only are often displayed only when pivot/table is refreshed.
/// Excel doesn't recalculate border styles for pivot table, unless necessary.
/// </remarks>
internal record XlsxDifferentialBorderValue
{
    internal static readonly XlsxDifferentialBorderValue Empty = new()
    {
        Left = null,
        Right = null,
        Top = null,
        Bottom = null,
        Diagonal = null,
        Vertical = null,
        Horizontal = null,
        DiagonalUp = false,
        DiagonalDown = false,
        Outline = true,
    };

    public required XlsxBorderLine? Left { get; init; }

    public required XlsxBorderLine? Right { get; init; }

    public required XlsxBorderLine? Top { get; init; }

    public required XlsxBorderLine? Bottom { get; init; }

    /// <summary>
    /// For pivot/tables only. Doesn't work on CF. Used when <see cref="DiagonalUp"/> or
    /// <see cref="DiagonalDown"/> are set. It's not possible to have different style
    /// for up/down diagonal.
    /// </summary>
    public required XlsxBorderLine? Diagonal { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required XlsxBorderLine? Vertical { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required XlsxBorderLine? Horizontal { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required bool DiagonalUp { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required bool DiagonalDown { get; init; }

    /// <summary>
    /// For pivot/tables or tables only.
    /// </summary>
    public required bool Outline { get; init; }
}
