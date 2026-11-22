using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles.Formatting;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.RichText
{
    internal class XlsxPhonetics : IXlsxPhonetics
    {
        private readonly List<IXlsxPhonetic> _phonetics = new();
        private readonly XlsxWorkbookStyles _styles;
        private readonly XlsxFontFormatValue _defaultFont;
        private readonly Action _onChange;
        private XlsxFontFormatValue _font;
        private XlsxPhoneticAlignment _alignment;
        private XlsxPhoneticType _type;

        public XlsxPhonetics(XlsxFontFormatValue font, XlsxFontFormatValue defaultFont, XlsxWorkbookStyles styles, Action onChange)
        {
            _styles = styles;
            _defaultFont = defaultFont;
            _font = font;
            _type = XlsxPhoneticType.FullWidthKatakana;
            _alignment = XlsxPhoneticAlignment.Left;
            _onChange = onChange;
        }

        public Int32 Count => _phonetics.Count;

        public Boolean Bold
        {
            get => _font.Bold;
            set => ChangeFont(f => f with { Bold = value });
        }

        public Boolean Italic
        {
            get => _font.Italic;
            set => ChangeFont(f => f with { Italic = value });
        }

        public XlsxFontUnderlineValues Underline
        {
            get => _font.Underline;
            set => ChangeFont(f => f with { Underline = value });
        }

        public Boolean Strikethrough
        {
            get => _font.Strikethrough;
            set => ChangeFont(f => f with { Strikethrough = value });
        }

        public XlsxFontVerticalTextAlignmentValues VerticalAlignment
        {
            get => _font.VerticalAlignment;
            set => ChangeFont(f => f with { VerticalAlignment = value });
        }

        public Boolean Shadow
        {
            get => _font.Shadow;
            set => ChangeFont(f => f with { Shadow = value });
        }

        public Double FontSize
        {
            get => _font.Size.Points;
            set => ChangeFont(f => f with { Size = XlsxFontSize.FromPoints(value) });
        }

        public XlsxColor FontColor
        {
            get => _font.Color;
            set => ChangeFont(f => f with { Color = value });
        }

        public String FontName
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

        public XlsxPhoneticAlignment Alignment
        {
            get => _alignment;
            set
            {
                _alignment = value;
                _onChange();
            }
        }

        public XlsxPhoneticType Type
        {
            get => _type;
            set
            {
                _type = value;
                _onChange();
            }
        }

        internal XlsxFontFormatValue Font => _font;

        public IXlsxPhonetics SetBold() { Bold = true; return this; }

        public IXlsxPhonetics SetBold(Boolean value) { Bold = value; return this; }

        public IXlsxPhonetics SetItalic() { Italic = true; return this; }

        public IXlsxPhonetics SetItalic(Boolean value) { Italic = value; return this; }

        public IXlsxPhonetics SetUnderline() { Underline = XlsxFontUnderlineValues.Single; return this; }

        public IXlsxPhonetics SetUnderline(XlsxFontUnderlineValues value) { Underline = value; return this; }

        public IXlsxPhonetics SetStrikethrough() { Strikethrough = true; return this; }

        public IXlsxPhonetics SetStrikethrough(Boolean value) { Strikethrough = value; return this; }

        public IXlsxPhonetics SetVerticalAlignment(XlsxFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return this; }

        public IXlsxPhonetics SetShadow() { Shadow = true; return this; }

        public IXlsxPhonetics SetShadow(Boolean value) { Shadow = value; return this; }

        public IXlsxPhonetics SetFontSize(Double value) { FontSize = value; return this; }

        public IXlsxPhonetics SetFontColor(XlsxColor value) { FontColor = value; return this; }

        public IXlsxPhonetics SetFontName(String value) { FontName = value; return this; }

        public IXlsxPhonetics SetFontFamilyNumbering(XlsxFontFamilyNumberingValues value) { FontFamilyNumbering = value; return this; }

        public IXlsxPhonetics SetFontCharSet(XlsxFontCharSet value) { FontCharSet = value; return this; }

        public IXlsxPhonetics SetFontScheme(XlsxFontScheme value) { FontScheme = value; return this; }

        public IXlsxPhonetics SetAlignment(XlsxPhoneticAlignment phoneticAlignment) { Alignment = phoneticAlignment; return this; }

        public IXlsxPhonetics SetType(XlsxPhoneticType phoneticType) { Type = phoneticType; return this; }

        public IXlsxPhonetics Add(String text, Int32 start, Int32 end)
        {
            _phonetics.Add(new XlsxPhonetic(text, start, end));
            _onChange();
            return this;
        }

        public IXlsxPhonetics ClearText()
        {
            _phonetics.Clear();
            _onChange();
            return this;
        }

        public IXlsxPhonetics ClearFont()
        {
            _font = _defaultFont;
            _onChange();
            return this;
        }

        public IEnumerator<IXlsxPhonetic> GetEnumerator()
        {
            return _phonetics.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool Equals(IXlsxPhonetics? other) => Equals(other as XlsxPhonetics);

        public bool Equals(XlsxPhonetics? other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            if (!_phonetics.SequenceEqual(other._phonetics))
                return false;

            return
                _font.Equals(other._font) &&
                Type == other.Type &&
                Alignment == other.Alignment;
        }

        private void ChangeFont(Func<XlsxFontFormatValue, XlsxFontFormatValue> modifyFont)
        {
            _font = _styles.GetRegisteredFontFormat(_font, modifyFont);
            _onChange();
        }
    }
}
