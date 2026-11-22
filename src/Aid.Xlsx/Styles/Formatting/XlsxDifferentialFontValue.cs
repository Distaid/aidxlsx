using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// A differential font format.
/// </summary>
internal record XlsxDifferentialFontValue
{
    /// <summary>
    /// A value with all properties null.
    /// </summary>
    internal static XlsxDifferentialFontValue Empty = new()
    {
        Name = null,
        Charset = null,
        Family = null,
        Bold = null,
        Italic = null,
        Strikethrough = null,
        Outline = null,
        Shadow = null,
        Condense = null,
        Extend = null,
        Color = null,
        Size = null,
        Underline = null,
        VerticalAlignment = null,
        Scheme = null,
    };

    internal required XlsxFontName? Name { get; init; }

    internal required XlsxFontCharSet? Charset { get; init; }

    internal required XlsxFontFamilyNumberingValues? Family { get; init; }

    internal required bool? Bold { get; init; }

    internal required bool? Italic { get; init; }

    internal required bool? Strikethrough { get; init; }

    internal required bool? Outline { get; init; }

    internal required bool? Shadow { get; init; }

    internal required bool? Condense { get; init; }

    internal required bool? Extend { get; init; }

    internal required XlsxColor? Color { get; init; }

    internal required XlsxFontSize? Size { get; init; }

    internal required XlsxFontUnderlineValues? Underline { get; init; }

    internal required XlsxFontVerticalTextAlignmentValues? VerticalAlignment { get; init; }

    internal required XlsxFontScheme? Scheme { get; init; }

    internal bool IsEmpty()
    {
        return Name is null &&
               Charset is null &&
               Family is null &&
               Bold is null &&
               Italic is null &&
               Strikethrough is null &&
               Outline is null &&
               Shadow is null &&
               Condense is null &&
               Extend is null &&
               Color is null &&
               Size is null &&
               Underline is null &&
               VerticalAlignment is null &&
               Scheme is null;
    }
}
