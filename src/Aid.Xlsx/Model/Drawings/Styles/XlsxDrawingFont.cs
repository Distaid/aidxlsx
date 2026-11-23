#nullable disable

using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingFont(IXlsxDrawingStyle style) : IXlsxDrawingFont
{
    private readonly IXlsxDrawingStyle _style = style;

    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public XlsxFontUnderlineValues Underline { get; set; } = XlsxFontUnderlineValues.None;
    public bool Strikethrough { get; set; }
    public XlsxFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
    public bool Shadow { get; set; }
    public double FontSize { get; set; } = 9;
    public XlsxColor FontColor { get; set; } = XlsxColor.FromIndex(64);
    public string FontName { get; set; } = "Tahoma";
    public XlsxFontFamilyNumberingValues FontFamilyNumbering { get; set; }
    public XlsxFontCharSet FontCharSet { get; set; }
    public XlsxFontScheme FontScheme { get; set; }

    public IXlsxDrawingStyle SetBold()
    {
        Bold = true;
        return _style;
    }

    public IXlsxDrawingStyle SetBold(bool value)
    {
        Bold = value;
        return _style;
    }

    public IXlsxDrawingStyle SetItalic()
    {
        Italic = true;
        return _style;
    }

    public IXlsxDrawingStyle SetItalic(bool value)
    {
        Italic = value;
        return _style;
    }

    public IXlsxDrawingStyle SetUnderline()
    {
        Underline = XlsxFontUnderlineValues.Single;
        return _style;
    }

    public IXlsxDrawingStyle SetUnderline(XlsxFontUnderlineValues value)
    {
        Underline = value;
        return _style;
    }

    public IXlsxDrawingStyle SetStrikethrough()
    {
        Strikethrough = true;
        return _style;
    }

    public IXlsxDrawingStyle SetStrikethrough(bool value)
    {
        Strikethrough = value;
        return _style;
    }

    public IXlsxDrawingStyle SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value)
    {
        VerticalAlignment = value;
        return _style;
    }

    public IXlsxDrawingStyle SetShadow()
    {
        Shadow = true;
        return _style;
    }

    public IXlsxDrawingStyle SetShadow(bool value)
    {
        Shadow = value;
        return _style;
    }

    public IXlsxDrawingStyle SetFontSize(double value)
    {
        FontSize = value;
        return _style;
    }

    public IXlsxDrawingStyle SetFontColor(XlsxColor value)
    {
        FontColor = value;
        return _style;
    }

    public IXlsxDrawingStyle SetFontName(string value)
    {
        FontName = value;
        return _style;
    }

    public IXlsxDrawingStyle SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value)
    {
        FontFamilyNumbering = value;
        return _style;
    }

    public IXlsxDrawingStyle SetFontCharSet(XlsxFontCharSet value)
    {
        FontCharSet = value;
        return _style;
    }

    public IXlsxDrawingStyle SetFontScheme(XlsxFontScheme value)
    {
        FontScheme = value;
        return _style;
    }
}
