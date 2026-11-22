using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

internal partial class XlsxDxfFontFormat : IXlsxFont
{
    private readonly XlsxFontFormatValue _defaultFont = XlsxFontFormatValue.Default;

    bool IXlsxFontBase.Bold
    {
        get => Resolve(static font => font.Bold, _defaultFont.Bold);
        set => Modify(static (font, bold) => font with { Bold = bold }, value);
    }

    bool IXlsxFontBase.Italic
    {
        get => Resolve(static font => font.Italic, _defaultFont.Italic);
        set => Modify(static (font, italic) => font with { Italic = italic }, value);
    }

    XlsxFontUnderlineValues IXlsxFontBase.Underline
    {
        get => Resolve(static font => font.Underline, _defaultFont.Underline);
        set => Modify(static (font, underline) => font with { Underline = underline }, value);
    }

    bool IXlsxFontBase.Strikethrough
    {
        get => Resolve(static font => font.Strikethrough, _defaultFont.Strikethrough);
        set => Modify(static (font, strikethrough) => font with { Strikethrough = strikethrough }, value);
    }

    XlsxFontVerticalTextAlignmentValues IXlsxFontBase.VerticalAlignment
    {
        get => Resolve(static font => font.VerticalAlignment, _defaultFont.VerticalAlignment);
        set => Modify(static (font, vAlign) => font with { VerticalAlignment = vAlign }, value);
    }

    bool IXlsxFontBase.Shadow
    {
        get => Resolve(static font => font.Shadow, _defaultFont.Shadow);
        set => Modify(static (font, shadow) => font with { Shadow = shadow }, value);
    }

    double IXlsxFontBase.FontSize
    {
        get => Resolve(static font => font.Size?.Points, _defaultFont.Size.Points);
        set => Modify(static (font, size) => font with { Size = XlsxFontSize.FromPoints(size) }, value);
    }

    XlsxColor IXlsxFontBase.FontColor
    {
        get => Resolve(static font => font.Color, _defaultFont.Color);
        set => Modify(static (font, color) => font with { Color = color }, value);
    }

    string IXlsxFontBase.FontName
    {
        get => Resolve(static font => font.Name?.Text, _defaultFont.Name.Text);
        set => Modify(static (font, name) => font with { Name = name }, value);
    }

    XlsxFontFamilyNumberingValues IXlsxFontBase.FontFamilyNumbering
    {
        get => Resolve(static font => font.Family, _defaultFont.Family);
        set => Modify(static (font, family) => font with { Family = family }, value);
    }

    XlsxFontCharSet IXlsxFontBase.FontCharSet
    {
        get => Resolve(static font => font.Charset, _defaultFont.Charset);
        set => Modify(static (font, charset) => font with { Charset = charset }, value);
    }

    XlsxFontScheme IXlsxFontBase.FontScheme
    {
        get => Resolve(static font => font.Scheme, _defaultFont.Scheme);
        set => Modify(static (font, scheme) => font with { Scheme = scheme }, value);
    }

    IXlsxStyle IXlsxFont.SetBold()
    {
        return (this as IXlsxFont).SetBold(true);
    }

    IXlsxStyle IXlsxFont.SetBold(bool value)
    {
        (this as IXlsxFont).Bold = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetItalic()
    {
        return (this as IXlsxFont).SetItalic(true);
    }

    IXlsxStyle IXlsxFont.SetItalic(bool value)
    {
        (this as IXlsxFont).Italic = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetUnderline()
    {
        return (this as IXlsxFont).SetUnderline(XlsxFontUnderlineValues.Single);
    }

    IXlsxStyle IXlsxFont.SetUnderline(XlsxFontUnderlineValues value)
    {
        (this as IXlsxFont).Underline = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetStrikethrough()
    {
        return (this as IXlsxFont).SetStrikethrough(true);
    }

    IXlsxStyle IXlsxFont.SetStrikethrough(bool value)
    {
        (this as IXlsxFont).Strikethrough = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value)
    {
        (this as IXlsxFont).VerticalAlignment = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetShadow()
    {
        return (this as IXlsxFont).SetShadow(true);
    }

    IXlsxStyle IXlsxFont.SetShadow(bool value)
    {
        (this as IXlsxFont).Shadow = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontSize(double value)
    {
        (this as IXlsxFont).FontSize = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontColor(XlsxColor value)
    {
        (this as IXlsxFont).FontColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontName(string value)
    {
        (this as IXlsxFont).FontName = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value)
    {
        (this as IXlsxFont).FontFamilyNumbering = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontCharSet(XlsxFontCharSet value)
    {
        (this as IXlsxFont).FontCharSet = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontScheme(XlsxFontScheme value)
    {
        (this as IXlsxFont).FontScheme = value;
        return _parent;
    }

    bool IEquatable<IXlsxFont>.Equals(IXlsxFont? other)
    {
        throw new NotSupportedException();
    }
}
