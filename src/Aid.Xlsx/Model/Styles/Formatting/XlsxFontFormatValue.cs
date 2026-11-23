using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles.Formatting;

/// <summary>
/// A formatting record for <see cref="XlsxCellFormatValue"/>.
/// </summary>
internal record XlsxFontFormatValue
{
    /// <summary>
    /// <para>
    /// A default font values.
    /// <list type="bullet">
    ///   <item>
    ///   When a font is loaded from the styles part, unspecified props use these values. The values
    ///   are essentially reinterpreted <em>zero</em> (other than name and size that can't be zero).
    ///   </item>
    ///   <item>
    ///   When font property has this value, it is omitted from being saved to font element (other
    ///   than font and size). Font element is pretty buggy, but these are roughly XSD default values.
    ///   </item>
    /// </list>
    /// The default values are different from normal style. Normal style is what should default
    /// format look like for user when a new workbook is created. The default font is about loading and
    /// saving.
    /// </para>
    /// <para>
    /// The font name and size are special, but the values in the property are the default fallback
    /// when even default format doesn't contain font name and size.
    /// </para>
    /// </summary>
    public static readonly XlsxFontFormatValue Default = new()
    {
        Name = "Calibri",
        Charset = XlsxFontCharSet.Ansi,
        Family = XlsxFontFamilyNumberingValues.NotApplicable,
        Bold = false,
        Italic = false,
        Strikethrough = false,
        Outline = false,
        Shadow = false,
        Condense = false,
        Extend = false,
        Color = XlsxColor.Automatic,
        Size = XlsxFontSize.FromPoints(11),
        Underline = XlsxFontUnderlineValues.None,
        VerticalAlignment = XlsxFontVerticalTextAlignmentValues.Baseline,
        Scheme = XlsxFontScheme.None
    };

    public required XlsxFontName Name { get; init; }

    public required XlsxFontCharSet Charset { get; init; }

    public required XlsxFontFamilyNumberingValues Family { get; init; }

    public required bool Bold { get; init; }

    public required bool Italic { get; init; }

    public required bool Strikethrough { get; init; }

    public required bool Outline { get; init; }

    public required bool Shadow { get; init; }

    public required bool Condense { get; init; }

    public required bool Extend { get; init; }

    public required XlsxColor Color { get; init; }

    public required XlsxFontSize Size { get; init; }

    public required XlsxFontUnderlineValues Underline { get; init; }

    public required XlsxFontVerticalTextAlignmentValues VerticalAlignment { get; init; }

    public required XlsxFontScheme Scheme { get; init; }

    /// <summary>
    /// Return a registered font format equivalent to <paramref name="font"/>.
    /// </summary>
    internal static XlsxFontFormatValue FromFontBase(IXlsxFontBase font, XlsxWorkbookStyles styles)
    {
        var fontFormat = new XlsxFontFormatValue
        {
            Name = font.FontName,
            Charset = font.FontCharSet,
            Family = font.FontFamilyNumbering,
            Bold = font.Bold,
            Italic = font.Italic,
            Strikethrough = font.Strikethrough,
            Outline = Default.Outline,
            Shadow = font.Shadow,
            Condense = Default.Condense,
            Extend = Default.Extend,
            Color = font.FontColor,
            Size = XlsxFontSize.FromPoints(font.FontSize),
            Underline = font.Underline,
            VerticalAlignment = font.VerticalAlignment,
            Scheme = font.FontScheme
        };
        return styles.GetRegisteredFontFormat(fontFormat, static x => x);
    }

    /// <summary>
    /// Create an adapter to font base. The adapter is not modifiable.
    /// </summary>
    internal IXlsxFontBase ToFontBase()
    {
        return new FontBaseAdapter(this);
    }

    /// <summary>
    /// Return a font that will have its values taken from the dxf where the dxf does have value.
    /// </summary>
    internal XlsxFontFormatValue AdjustWith(XlsxDifferentialFontValue dxf)
    {
        if (dxf.IsEmpty())
            return this;

        XlsxFontFormatValue result = this;
        if (dxf.Name is { } name)
            result = result with { Name = name };

        if (dxf.Charset is { } charset)
            result = result with { Charset = charset };

        if (dxf.Family is { } family)
            result = result with { Family = family };

        if (dxf.Bold is { } bold)
            result = result with { Bold = bold };

        if (dxf.Italic is { } italic)
            result = result with { Italic = italic };

        if (dxf.Strikethrough is { } strikethrough)
            result = result with { Strikethrough = strikethrough };

        if (dxf.Outline is { } outline)
            result = result with { Outline = outline };

        if (dxf.Shadow is { } shadow)
            result = result with { Shadow = shadow };

        if (dxf.Condense is { } condense)
            result = result with { Condense = condense };

        if (dxf.Extend is { } extend)
            result = result with { Extend = extend };

        if (dxf.Color is { } color)
            result = result with { Color = color };

        if (dxf.Size is { } size)
            result = result with { Size = size };

        if (dxf.Underline is { } underline)
            result = result with { Underline = underline };

        if (dxf.VerticalAlignment is { } verticalAlignment)
            result = result with { VerticalAlignment = verticalAlignment };

        if (dxf.Scheme is { } scheme)
            result = result with { Scheme = scheme };

        return result;
    }

    private class FontBaseAdapter : IXlsxFontBase
    {
        private readonly XlsxFontFormatValue _font;

        internal FontBaseAdapter(XlsxFontFormatValue font)
        {
            _font = font;
        }

        public bool Bold
        {
            get => _font.Bold;
            set => throw Exception();
        }

        public bool Italic
        {
            get => _font.Italic;
            set => throw Exception();
        }
        public XlsxFontUnderlineValues Underline
        {
            get => _font.Underline;
            set => throw Exception();
        }

        public bool Strikethrough
        {
            get => _font.Strikethrough;
            set => throw Exception();
        }

        public XlsxFontVerticalTextAlignmentValues VerticalAlignment
        {
            get => _font.VerticalAlignment;
            set => throw Exception();
        }

        public bool Shadow
        {
            get => _font.Shadow;
            set => throw Exception();
        }

        public double FontSize
        {
            get => _font.Size.Points;
            set => throw Exception();
        }

        public XlsxColor FontColor
        {
            get => _font.Color;
            set => throw Exception();
        }

        public string FontName
        {
            get => _font.Name.Text;
            set => throw Exception();
        }

        public XlsxFontFamilyNumberingValues FontFamilyNumbering
        {
            get => _font.Family;
            set => throw Exception();
        }

        public XlsxFontCharSet FontCharSet
        {
            get => _font.Charset;
            set => throw Exception();
        }

        public XlsxFontScheme FontScheme
        {
            get => _font.Scheme;
            set => throw Exception();
        }

        private NotSupportedException Exception()
        {
            return new NotSupportedException("This is an adapter for an immutable font format.");
        }
    }
}
