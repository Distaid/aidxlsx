using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXlsxFormatContainer"/>.
/// </summary>
internal sealed partial class XlsxFontCellFormat
{
    private readonly XlsxCellFormat _parent;

    internal XlsxFontCellFormat(XlsxCellFormat parent)
    {
        _parent = parent;
    }

    internal XlsxFontName Name
    {
        get => Resolve(static x => x.Font.Name);
        set => Modify(static (font, fontName) => font with { Name = fontName }, value);
    }

    internal XlsxFontCharSet Charset
    {
        get => Resolve(static x => x.Font.Charset);
        set => Modify(static (font, charset) => font with { Charset = charset }, value);
    }

    internal XlsxFontFamilyNumberingValues Family
    {
        get => Resolve(static x => x.Font.Family);
        set => Modify(static (font, family) => font with { Family = family }, value);
    }

    internal bool Bold
    {
        get => Resolve(static x => x.Font.Bold);
        set => Modify(static (font, bold) => font with { Bold = bold }, value);
    }

    internal bool Italic
    {
        get => Resolve(static x => x.Font.Italic);
        set => Modify(static (font, italic) => font with { Italic = italic }, value);
    }

    internal bool Strikethrough
    {
        get => Resolve(static x => x.Font.Strikethrough);
        set => Modify(static (font, strikethrough) => font with { Strikethrough = strikethrough }, value);
    }

    internal bool Outline
    {
        get => Resolve(static x => x.Font.Outline);
        set => Modify(static (font, outline) => font with { Outline = outline }, value);
    }

    internal bool Shadow
    {
        get => Resolve(static x => x.Font.Shadow);
        set => Modify(static (font, shadow) => font with { Shadow = shadow }, value);
    }

    internal XlsxColor Color
    {
        get => Resolve(static x => x.Font.Color);
        set => Modify(static (font, color) => font with { Color = color }, value);
    }

    internal XlsxFontSize Size
    {
        get => Resolve(static x => x.Font.Size);
        set => Modify(static (font, size) => font with { Size = size }, value);
    }

    internal XlsxFontUnderlineValues Underline
    {
        get => Resolve(static x => x.Font.Underline);
        set => Modify(static (font, underline) => font with { Underline = underline }, value);
    }

    internal XlsxFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => Resolve(static x => x.Font.VerticalAlignment);
        set => Modify(static (font, verticalAlignment) => font with { VerticalAlignment = verticalAlignment }, value);
    }

    internal XlsxFontScheme Scheme
    {
        get => Resolve(static x => x.Font.Scheme);
        set => Modify(static (font, scheme) => font with { Scheme = scheme }, value);
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxFont other && (this as IEquatable<IXlsxFont>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    private T Resolve<T>(Func<XlsxCellFormatValue, T> selector)
    {
        return _parent.Resolve(selector);
    }

    private void Modify<TProperty>(Func<XlsxFontFormatValue, TProperty, XlsxFontFormatValue> modifyFont, TProperty value)
    {
        _parent.ModifyFont(modifyFont, value);
    }

    /// <summary>
    /// A helper method to set all font properties at once (e.g, <c>someStyle.Font = otherStyle.Font</c>).
    /// </summary>
    internal void SetFont(IXlsxFont value)
    {
        _parent.ModifyFont(static (font, value) => font with
        {
            Bold = value.Bold,
            Italic = value.Italic,
            Underline = value.Underline,
            Strikethrough = value.Strikethrough,
            VerticalAlignment = value.VerticalAlignment,
            Shadow = value.Shadow,
            Size = XlsxFontSize.FromPoints(value.FontSize),
            Color = value.FontColor,
            Name = value.FontName,
            Family = value.FontFamilyNumbering,
            Charset = value.FontCharSet,
            Scheme = value.FontScheme
        }, value);
    }
}
