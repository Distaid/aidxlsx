using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Styles;

internal sealed partial class XlsxFontCellFormat : IXlsxFont
{
    bool IXlsxFontBase.Bold
    {
        get => Bold;
        set => Bold = value;
    }

    bool IXlsxFontBase.Italic
    {
        get => Italic;
        set => Italic = value;
    }

    XlsxFontUnderlineValues IXlsxFontBase.Underline
    {
        get => Underline;
        set => Underline = value;
    }

    bool IXlsxFontBase.Strikethrough
    {
        get => Strikethrough;
        set => Strikethrough = value;
    }

    XlsxFontVerticalTextAlignmentValues IXlsxFontBase.VerticalAlignment
    {
        get => VerticalAlignment;
        set => VerticalAlignment = value;
    }

    bool IXlsxFontBase.Shadow
    {
        get => Shadow;
        set => Shadow = value;
    }

    double IXlsxFontBase.FontSize
    {
        get => Size.Points;
        set => Size = XlsxFontSize.FromPoints(value);
    }

    XlsxColor IXlsxFontBase.FontColor
    {
        get => Color;
        set => Color = value;
    }

    string IXlsxFontBase.FontName
    {
        get => Name.Text;
        set => Name = value;
    }

    XlsxFontFamilyNumberingValues IXlsxFontBase.FontFamilyNumbering
    {
        get => Family;
        set => Family = value;
    }

    XlsxFontCharSet IXlsxFontBase.FontCharSet
    {
        get => Charset;
        set => Charset = value;
    }

    XlsxFontScheme IXlsxFontBase.FontScheme
    {
        get => Scheme;
        set => Scheme = value;
    }

    bool IEquatable<IXlsxFont>.Equals(IXlsxFont? other)
    {
        if (other is null)
            return false;

        if (Bold != other.Bold)
            return false;

        if (Italic != other.Italic)
            return false;

        if (Underline != other.Underline)
            return false;

        if (Strikethrough != other.Strikethrough)
            return false;

        if (VerticalAlignment != other.VerticalAlignment)
            return false;

        if (Shadow != other.Shadow)
            return false;

        if (!Size.Points.Equals(other.FontSize))
            return false;

        if (Color != other.FontColor)
            return false;

        if (Name != other.FontName)
            return false;

        if (Family != other.FontFamilyNumbering)
            return false;

        if (Charset != other.FontCharSet)
            return false;

        if (Scheme != other.FontScheme)
            return false;

        return true;
    }

    IXlsxStyle IXlsxFont.SetBold()
    {
        return (this as IXlsxFont).SetBold(true);
    }

    IXlsxStyle IXlsxFont.SetBold(bool value)
    {
        Bold = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetItalic()
    {
        return (this as IXlsxFont).SetItalic(true);
    }

    IXlsxStyle IXlsxFont.SetItalic(bool value)
    {
        Italic = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetUnderline()
    {
        return (this as IXlsxFont).SetUnderline(XlsxFontUnderlineValues.Single);
    }

    IXlsxStyle IXlsxFont.SetUnderline(XlsxFontUnderlineValues value)
    {
        Underline = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetStrikethrough()
    {
        return (this as IXlsxFont).SetStrikethrough(true);
    }

    IXlsxStyle IXlsxFont.SetStrikethrough(bool value)
    {
        Strikethrough = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value)
    {
        VerticalAlignment = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetShadow()
    {
        return (this as IXlsxFont).SetShadow(true);
    }

    IXlsxStyle IXlsxFont.SetShadow(bool value)
    {
        Shadow = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontSize(double value)
    {
        (this as IXlsxFont).FontSize = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontColor(XlsxColor value)
    {
        Color = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontName(string value)
    {
        Name = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value)
    {
        Family = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontCharSet(XlsxFontCharSet value)
    {
        Charset = value;
        return _parent;
    }

    IXlsxStyle IXlsxFont.SetFontScheme(XlsxFontScheme value)
    {
        Scheme = value;
        return _parent;
    }
}
