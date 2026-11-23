#nullable disable

using System.Diagnostics;
using System.Text;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.RichText;

internal class XlsxFormattedText<T> : IXlsxFormattedText<T>
{
    // TODO: Move from ancestor to children, not needed here
    protected T Container;
    protected readonly XlsxWorkbookStyles Styles;

    /// <summary>
    /// Font used for a new rich text run, never modified. It is generally provided by a container of the formatted text.
    /// </summary>
    private readonly XlsxFontFormatValue _defaultFont;
    private readonly List<XlsxRichString> _richTexts = [];
    private XlsxPhonetics _phonetics;

    protected XlsxFormattedText(XlsxFontFormatValue defaultFont, XlsxWorkbookStyles styles)
    {
        Debug.Assert(styles.Fonts.ContainsValue(defaultFont));
        _defaultFont = defaultFont;
        Styles = styles;
    }

    IXlsxPhonetics IXlsxFormattedText<T>.Phonetics => Phonetics;

    public int Count => _richTexts.Count;

    public int Length { get; private set; }

    public string Text => ToString();

    public bool HasPhonetics => _phonetics is not null;

    /// <inheritdoc cref="IXlsxFormattedText{T}.Phonetics"/>
    internal XlsxPhonetics Phonetics
    {
        get => _phonetics ??= new XlsxPhonetics(_defaultFont, _defaultFont, Styles, OnContentChanged);
        init => _phonetics = value;
    }

    public IXlsxRichString AddText(string text)
    {
        return AddText(text, _defaultFont);
    }

    public IXlsxRichString AddText(string text, IXlsxFontBase font)
    {
        return AddText(text, XlsxFontFormatValue.FromFontBase(font, Styles));
    }

    internal void AddText(string text, XlsxDifferentialFontValue font)
    {
        AddText(text, _defaultFont.AdjustWith(font));
    }

    private IXlsxRichString AddText(string text, XlsxFontFormatValue font)
    {
        var richText = new XlsxRichString(text, font, this, Styles, OnContentChanged);
        AddText(richText);
        OnContentChanged();
        return richText;
    }

    public IXlsxRichString AddNewLine()
    {
        return AddText(Environment.NewLine);
    }

    public IXlsxFormattedText<T> ClearText()
    {
        ClearContent();
        OnContentChanged();
        return this;
    }

    public IXlsxFormattedText<T> ClearFont()
    {
        var text = Text;
        ClearContent();
        AddText(text);
        return this;
    }

    public override string ToString()
    {
        var sb = new StringBuilder(_richTexts.Count);
        _richTexts.ForEach(rt => sb.Append(rt.Text));
        return sb.ToString();
    }

    public IXlsxFormattedText<T> Substring(int index)
    {
        return Substring(index, Length - index);
    }

    public IXlsxFormattedText<T> Substring(int index, int length)
    {
        if (index + 1 > Length || (Length - index + 1) < length || length <= 0)
            throw new IndexOutOfRangeException("Index and length must refer to a location within the string.");

        var newRichTexts = new List<XlsxRichString>();
        var retVal = new XlsxFormattedText<T>(_defaultFont, Styles);

        var lastPosition = 0;
        foreach (XlsxRichString rt in _richTexts)
        {
            if (lastPosition >= index + 1 + length) // We already have what we need
            {
                newRichTexts.Add(rt);
            }
            else if (lastPosition + rt.Text.Length >= index + 1) // Eureka!
            {
                var startIndex = index - lastPosition;

                if (startIndex > 0)
                    newRichTexts.Add(new XlsxRichString(rt.Text[..startIndex], rt.Font, this, Styles, OnContentChanged));
                else if (startIndex < 0)
                    startIndex = 0;

                var leftToTake = length - retVal.Length;
                if (leftToTake > rt.Text.Length - startIndex)
                    leftToTake = rt.Text.Length - startIndex;

                var newRt = new XlsxRichString(rt.Text.Substring(startIndex, leftToTake), rt.Font, this, Styles, OnContentChanged);
                newRichTexts.Add(newRt);
                retVal.AddText(newRt);

                if (startIndex + leftToTake < rt.Text.Length)
                    newRichTexts.Add(new XlsxRichString(rt.Text[(startIndex + leftToTake)..], rt.Font, this, Styles, OnContentChanged));
            }
            else // We haven't reached the desired position yet
            {
                newRichTexts.Add(rt);
            }
            lastPosition += rt.Text.Length;
        }

        _richTexts.Clear();
        _richTexts.AddRange(newRichTexts);
        OnContentChanged();
        return retVal;
    }

    public IXlsxFormattedText<T> CopyFrom(IXlsxFormattedText<T> original)
    {
        ClearContent();
        foreach (IXlsxRichString richText in original)
        {
            var copyFont = XlsxFontFormatValue.FromFontBase(richText, Styles);
            var copyText = new XlsxRichString(richText.Text, copyFont, this, Styles, OnContentChanged);
            AddText(copyText);
        }

        OnContentChanged();
        return this;
    }

    public List<XlsxRichString>.Enumerator GetEnumerator() => _richTexts.GetEnumerator();

    IEnumerator<IXlsxRichString> IEnumerable<IXlsxRichString>.GetEnumerator() => GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

    public bool Bold { set { _richTexts.ForEach(rt => rt.Bold = value); } }
    public bool Italic { set { _richTexts.ForEach(rt => rt.Italic = value); } }
    public XlsxFontUnderlineValues Underline { set { _richTexts.ForEach(rt => rt.Underline = value); } }
    public bool Strikethrough { set { _richTexts.ForEach(rt => rt.Strikethrough = value); } }
    public XlsxFontVerticalTextAlignmentValues VerticalAlignment { set { _richTexts.ForEach(rt => rt.VerticalAlignment = value); } }
    public bool Shadow { set { _richTexts.ForEach(rt => rt.Shadow = value); } }
    public double FontSize { set { _richTexts.ForEach(rt => rt.FontSize = value); } }
    public XlsxColor FontColor { set { _richTexts.ForEach(rt => rt.FontColor = value); } }
    public string FontName { set { _richTexts.ForEach(rt => rt.FontName = value); } }
    public XlsxFontFamilyNumberingValues FontFamilyNumbering { set { _richTexts.ForEach(rt => rt.FontFamilyNumbering = value); } }

    public IXlsxFormattedText<T> SetBold() { Bold = true; return this; }
    public IXlsxFormattedText<T> SetBold(bool value) { Bold = value; return this; }
    public IXlsxFormattedText<T> SetItalic() { Italic = true; return this; }
    public IXlsxFormattedText<T> SetItalic(bool value) { Italic = value; return this; }
    public IXlsxFormattedText<T> SetUnderline() { Underline = XlsxFontUnderlineValues.Single; return this; }
    public IXlsxFormattedText<T> SetUnderline(XlsxFontUnderlineValues value) { Underline = value; return this; }
    public IXlsxFormattedText<T> SetStrikethrough() { Strikethrough = true; return this; }
    public IXlsxFormattedText<T> SetStrikethrough(bool value) { Strikethrough = value; return this; }
    public IXlsxFormattedText<T> SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return this; }
    public IXlsxFormattedText<T> SetShadow() { Shadow = true; return this; }
    public IXlsxFormattedText<T> SetShadow(bool value) { Shadow = value; return this; }
    public IXlsxFormattedText<T> SetFontSize(double value) { FontSize = value; return this; }
    public IXlsxFormattedText<T> SetFontColor(XlsxColor value) { FontColor = value; return this; }
    public IXlsxFormattedText<T> SetFontName(string value) { FontName = value; return this; }
    public IXlsxFormattedText<T> SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value) { FontFamilyNumbering = value; return this; }

    public bool Equals(IXlsxFormattedText<T> other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        if (Count != other.Count)
            return false;

        if (!_richTexts.SequenceEqual(other))
            return false;

        return (_phonetics is null && !other.HasPhonetics) || Phonetics.Equals(other.Phonetics);
    }

    protected void AddText(XlsxRichString richText)
    {
        _richTexts.Add(richText);
        Length += richText.Text.Length;
    }

    /// <summary>
    /// This method is called every time the formatted text is changed (new runs, font props, phonetics...).
    /// </summary>
    protected virtual void OnContentChanged()
    {
        // Do nothing, intended to be overriden.
    }

    private void ClearContent()
    {
        _richTexts.Clear();
        Length = 0;
    }
}
