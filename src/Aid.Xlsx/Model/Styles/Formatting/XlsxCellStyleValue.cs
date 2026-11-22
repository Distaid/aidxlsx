namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// <para>
/// A cell style available in the workbook.
/// </para>
/// <para>
/// The style is intended to be mutable, unlike cell format. Changing a style aspect (name or size)
/// shouldn't require update of all cells that use it. The <see cref="XlsxCellFormatValue"/> links to
/// the style through <see cref="XlsxCellFormatValue.CellStyleId"/> and the <see cref="XlsxWorkbookStyles.CellStyles"/>
/// collection.
/// </para>
/// </summary>
internal class XlsxCellStyleValue
{
    /// <summary>
    /// A unique name of the style.
    /// </summary>
    public required string Name { get; init; }

    public required BuiltInStyleValues? BuiltInStyle { get; init; }

    /// <summary>
    /// Is style hidden in the UI?
    /// </summary>
    public required bool Hidden { get; init; }

    public required XlsxNumberFormat NumberFormat { get; init; }

    public required XlsxAlignmentFormatValue Alignment { get; init; }

    public required XlsxProtectionFormatValue Protection { get; init; }

    public required XlsxFontFormatValue Font { get; init; }

    public required XlsxFillFormatValue Fill { get; init; }

    public required XlsxBorderFormatValue Border { get; init; }

    /// <summary>
    /// Format components that are decided by the style. Specified components should have
    /// a non-null value.
    /// </summary>
    public required CellFormatComponents IncludedComponents { get; init; }
}
