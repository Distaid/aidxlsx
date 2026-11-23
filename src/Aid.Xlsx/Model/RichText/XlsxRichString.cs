using System.Diagnostics;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.RichText;

/// <summary>
/// An API object to modify a rich string.
/// </summary>
[DebuggerDisplay("{Text}")]
internal class XlsxRichString : IXlsxRichString
{
    private readonly XlsxWorkbookStyles _styles;
    private readonly IXlsxWithRichString _withRichString;
    private readonly Action _onChange;
    private string _text;
    private XlsxFontFormatValue _font;

    internal XlsxRichString(string text, XlsxFontFormatValue font, IXlsxWithRichString withRichString, XlsxWorkbookStyles styles, Action? onChange)
    {
        _text = text;
        _font = font;
        _withRichString = withRichString;
        _styles = styles;
        _onChange = onChange ?? (() => { });
    }

    public string Text
    {
        get => _text;
        set
        {
            _text = value;
            _onChange();
        }
    }

    public IXlsxRichString AddText(string text)
    {
        return _withRichString.AddText(text);
    }

    public IXlsxRichString AddNewLine()
    {
        return AddText(Environment.NewLine);
    }

    public bool Bold
    {
        get => _font.Bold;
        set => ChangeFont(f => f with { Bold = value });
    }

    public bool Italic
    {
        get => _font.Italic;
        set => ChangeFont(f => f with { Italic = value });
    }

    public XlsxFontUnderlineValues Underline
    {
        get => _font.Underline;
        set => ChangeFont(f => f with { Underline = value });
    }

    public bool Strikethrough
    {
        get => _font.Strikethrough;
        set => ChangeFont(f => f with { Strikethrough = value });
    }

    public XlsxFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => _font.VerticalAlignment;
        set => ChangeFont(f => f with { VerticalAlignment = value });
    }

    public bool Shadow
    {
        get => _font.Shadow;
        set => ChangeFont(f => f with { Shadow = value });
    }

    public double FontSize
    {
        get => _font.Size.Points;
        set => ChangeFont(f => f with { Size = XlsxFontSize.FromPoints(value) });
    }

    public XlsxColor FontColor
    {
        get => _font.Color;
        set => ChangeFont(f => f with { Color = value });
    }

    public string FontName
    {
        get => _font.Name.Text;
        set => ChangeFont(f => f with { Name = value });
    }

    public XlsxFontFamilyNumberingValues FontFamilyNumbering
    {
        get => _font.Family;
        set => ChangeFont(f => f with { Family = value });
    }

    public XlsxFontCharSet FontCharSet
    {
        get => _font.Charset;
        set => ChangeFont(f => f with { Charset = value });
    }

    public XlsxFontScheme FontScheme
    {
        get => _font.Scheme;
        set => ChangeFont(f => f with { Scheme = value });
    }

    internal XlsxFontFormatValue Font => _font;

    public IXlsxRichString SetBold()
    {
        Bold = true; return this;
    }

    public IXlsxRichString SetBold(bool value)
    {
        Bold = value; return this;
    }

    public IXlsxRichString SetItalic()
    {
        Italic = true; return this;
    }

    public IXlsxRichString SetItalic(bool value)
    {
        Italic = value; return this;
    }

    public IXlsxRichString SetUnderline()
    {
        Underline = XlsxFontUnderlineValues.Single; return this;
    }

    public IXlsxRichString SetUnderline(XlsxFontUnderlineValues value)
    {
        Underline = value; return this;
    }

    public IXlsxRichString SetStrikethrough()
    {
        Strikethrough = true; return this;
    }

    public IXlsxRichString SetStrikethrough(bool value)
    {
        Strikethrough = value; return this;
    }

    public IXlsxRichString SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value)
    {
        VerticalAlignment = value; return this;
    }

    public IXlsxRichString SetShadow()
    {
        Shadow = true; return this;
    }

    public IXlsxRichString SetShadow(bool value)
    {
        Shadow = value; return this;
    }

    public IXlsxRichString SetFontSize(double value)
    {
        FontSize = value; return this;
    }

    public IXlsxRichString SetFontColor(XlsxColor value)
    {
        FontColor = value; return this;
    }

    public IXlsxRichString SetFontName(string value)
    {
        FontName = value; return this;
    }

    public IXlsxRichString SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value)
    {
        FontFamilyNumbering = value; return this;
    }

    public IXlsxRichString SetFontCharSet(XlsxFontCharSet value)
    {
        FontCharSet = value; return this;
    }

    public IXlsxRichString SetFontScheme(XlsxFontScheme value)
    {
        FontScheme = value; return this;
    }

    public override bool Equals(object? obj) => Equals(obj as XlsxRichString);

    public bool Equals(IXlsxRichString? other) => Equals(other as XlsxRichString);

    public bool Equals(XlsxRichString? other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return Text == other.Text && _font.Equals(other._font);
    }

    public override int GetHashCode()
    {
        // Since all properties of type are mutable, can't have different hashcode for any instance.
        // Don't ever use this class in a dictionary, e.g. SST.
        return 4; // Chosen by fair dice roll. Guaranteed to be random.
    }

    private void ChangeFont(Func<XlsxFontFormatValue, XlsxFontFormatValue> modifyFont)
    {
        _font = _styles.GetRegisteredFontFormat(_font, modifyFont);
        _onChange();
    }
}
