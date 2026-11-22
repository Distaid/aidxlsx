namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// A differential format.
/// </summary>
internal record XlsxDxfValue
{
    internal static readonly XlsxDxfValue Empty = new()
    {
        NumberFormat = null,
        Font = XlsxDifferentialFontValue.Empty,
        Fill = XlsxDifferentialFillValue.Empty,
        Alignment = XlsxDifferentialAlignmentValue.Empty,
        Border = XlsxDifferentialBorderValue.Empty,
        Protection = XlsxDifferentialProtectionValue.Empty
    };

    public required string? NumberFormat { get; init; }

    public required XlsxDifferentialFontValue Font { get; init; }

    public required XlsxDifferentialFillValue Fill { get; init; }

    public required XlsxDifferentialAlignmentValue Alignment { get; init; }

    public required XlsxDifferentialBorderValue Border { get; init; }

    public required XlsxDifferentialProtectionValue Protection { get; init; }
}
