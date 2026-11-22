using System.Text;
using PhoneticProperties = Aid.Xlsx.Model.RichText.XlsxImmutableRichText.PhoneticProperties;
using PhoneticRun = Aid.Xlsx.Model.RichText.XlsxImmutableRichText.PhoneticRun;
using PhoneticRunDto = (string Text, int StartIndex, int EndIndex);
using RichTextRun = Aid.Xlsx.Model.RichText.XlsxImmutableRichText.RichTextRun;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Opc.Xml;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.SpreadsheetML.Strings;

/// <summary>
/// A partial stateful reader for reading a string item (it's either text or rich text). The rich
/// text format is not stored in the <see cref="XlsxWorkbookStyles"/> and the reader doesn't add it
/// there.
/// </summary>
internal partial class RstReader
{
    private static readonly Comparer<PhoneticRunDto> ComparePhoneticRunsByStartIndex = Comparer<PhoneticRunDto>.Create((x, y) => x.StartIndex.CompareTo(y.StartIndex));

    private readonly string _ns;

    private readonly XmlTreeReader _reader;

    private readonly XlsxWorkbookStyles _styles;

    // The values read from the CT_RPrElt element. The values are reset before each run.
    private XlsxFontName? _runName;
    private XlsxFontCharSet? _runCharset;
    private XlsxFontFamilyNumberingValues? _runFamily;
    private bool? _runBold;
    private bool? _runItalic;
    private bool? _runStrikethrough;
    private bool? _runOutline;
    private bool? _runShadow;
    private bool? _runCondense;
    private bool? _runExtend;
    private XlsxColor? _runColor;
    private XlsxFontSize? _runSize;
    private XlsxFontUnderlineValues? _runUnderline;
    private XlsxFontVerticalTextAlignmentValues? _runVerticalAlignment;
    private XlsxFontScheme? _runScheme;

    internal RstReader(XmlTreeReader reader, XlsxWorkbookStyles styles, string ns = OpenXmlConst.Main2006SsNs)
    {
        _reader = reader;
        _styles = styles;
        _ns = ns;
    }

    /// <summary>
    /// Parse a <c>CT_Rst</c> element.
    /// </summary>
    internal Xpr<OneOf<string, XlsxImmutableRichText>> ParseCtRst(string elementName, string ns)
    {
        ResetFontState();
        return ParseRst(elementName, ns);
    }

    /// <summary>
    /// Reset information about properties of current/next rich run.
    /// </summary>
    private void ResetFontState()
    {
        _runName = null;
        _runCharset = null;
        _runFamily = null;
        _runBold = null;
        _runItalic = null;
        _runStrikethrough = null;
        _runOutline = null;
        _runShadow = null;
        _runCondense = null;
        _runExtend = null;
        _runColor = null;
        _runSize = null;
        _runUnderline = null;
        _runVerticalAlignment = null;
        _runScheme = null;
    }

    /// <summary>
    /// Get the state of properties for current/next rich run.
    /// </summary>
    private XlsxDifferentialFontValue GetFontState()
    {
        return new XlsxDifferentialFontValue
        {
            Name = _runName,
            Charset = _runCharset,
            Family = _runFamily,
            Bold = _runBold,
            Italic = _runItalic,
            Strikethrough = _runStrikethrough,
            Outline = _runOutline,
            Shadow = _runShadow,
            Condense = _runCondense,
            Extend = _runExtend,
            Color = _runColor,
            Size = _runSize,
            Underline = _runUnderline,
            VerticalAlignment = _runVerticalAlignment,
            Scheme = _runScheme
        };
    }

    private bool OnBooleanPropertyParsed(bool value)
    {
        return value;
    }

    private int OnIntPropertyParsed(int value)
    {
        return value;
    }

    private XlsxFontSize OnFontSizeParsed(double sizePt)
    {
        return XlsxFontSize.FromPoints(sizePt);
    }

    private XlsxFontName OnFontNameParsed(string fontName)
    {
        return fontName;
    }

    private XlsxFontVerticalTextAlignmentValues OnVerticalAlignFontPropertyParsed(XlsxFontVerticalTextAlignmentValues verticalAlignment)
    {
        return verticalAlignment;
    }

    private XlsxFontScheme OnFontSchemeParsed(XlsxFontScheme scheme)
    {
        return scheme;
    }

    private XlsxFontUnderlineValues OnUnderlinePropertyParsed(XlsxFontUnderlineValues underline)
    {
        return underline;
    }


    #region Callbacks for the OnRPrElt that store the rich run properties

    private Unit OnRPrEltRFontParsed(XlsxFontName name)
    {
        _runName = name;
        return Unit.Value;
    }

    private Unit OnRPrEltCharsetParsed(int charsetValue)
    {
        if (charsetValue is < 0 or > 255)
            throw PartStructureException.InvalidAttributeValue(charsetValue.ToString());

        _runCharset = (XlsxFontCharSet)charsetValue;
        return Unit.Value;
    }

    private Unit OnRPrEltFamilyParsed(int familyValue)
    {
        // Unlike family in the CT_Font, the family in the CT_RPrElt is stored as an int.
        // Bug in the spec. Spec says that it has values 0-14 and doesn't specify meaning
        // for the numerical values. It's supposed to refer to the same enum ST_FontFamily
        // as in WordML. The OI-29500 fixes this problem:
        // "Excel restricts the value of this attribute to be at least 0 and at most 5."
        _runFamily = familyValue switch
        {
            >= 0 and <= 5 => (XlsxFontFamilyNumberingValues)familyValue,
            > 5 and <= 14 => XlsxFontFamilyNumberingValues.NotApplicable,
            _ => throw PartStructureException.InvalidAttributeValue(familyValue.ToString()),
        };
        return Unit.Value;
    }

    private Unit OnRPrEltBParsed(bool bold)
    {
        _runBold = bold;
        return Unit.Value;
    }

    private Unit OnRPrEltIParsed(bool italic)
    {
        _runItalic = italic;
        return Unit.Value;
    }

    private Unit OnRPrEltStrikeParsed(bool strikethrough)
    {
        _runStrikethrough = strikethrough;
        return Unit.Value;
    }

    private Unit OnRPrEltOutlineParsed(bool outline)
    {
        _runOutline = outline;
        return Unit.Value;
    }

    private Unit OnRPrEltShadowParsed(bool shadow)
    {
        _runShadow = shadow;
        return Unit.Value;
    }

    private Unit OnRPrEltCondenseParsed(bool condense)
    {
        _runCondense = condense;
        return Unit.Value;
    }

    private Unit OnRPrEltExtendParsed(bool extend)
    {
        _runExtend = extend;
        return Unit.Value;
    }

    private Xpr<XlsxColor> ParseColor(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxColor>();
        }

        return Xpr.From(_reader.ParseColor(elementName, ns));
    }

    private Unit OnRPrEltColorParsed(XlsxColor color)
    {
        _runColor = color;
        return Unit.Value;
    }

    private Unit OnRPrEltSzParsed(XlsxFontSize size)
    {
        _runSize = size;
        return Unit.Value;
    }

    private Unit OnRPrEltUParsed(XlsxFontUnderlineValues underline)
    {
        _runUnderline = underline;
        return Unit.Value;
    }

    private Unit OnRPrEltVertAlignParsed(XlsxFontVerticalTextAlignmentValues verticalAlignment)
    {
        _runVerticalAlignment = verticalAlignment;
        return Unit.Value;
    }

    private Unit OnRPrEltSchemeParsed(XlsxFontScheme scheme)
    {
        _runScheme = scheme;
        return Unit.Value;
    }

    /// <summary>
    /// Create font format from the values of the current state of the run (the <c>_run*</c> fields).
    /// </summary>
    private XlsxDifferentialFontValue OnRPrEltParsed(List<Unit> _)
    {
        return GetFontState();
    }

    #endregion

    private (string Text, XlsxDifferentialFontValue Font) OnREltParsed(XlsxDifferentialFontValue? runProperties, string runText)
    {
        var runFont = runProperties ?? XlsxDifferentialFontValue.Empty;
        ResetFontState();
        return (runText, runFont);
    }

    private PhoneticRunDto OnPhoneticRunParsed(string text, uint sb, uint eb)
    {
        // Validate and filter out later, in the method that construct the rich text
        return new PhoneticRunDto(text, checked((int)sb), checked((int)eb));
    }

    private PhoneticProperties OnPhoneticPrParsed(uint fontId, XlsxPhoneticType type, XlsxPhoneticAlignment alignment)
    {
        var phoneticFont = _styles.Fonts[checked((int)fontId)];
        return new PhoneticProperties(phoneticFont, type, alignment);
    }

    private OneOf<string, XlsxImmutableRichText> OnRstParsed(string? fullText, List<(string Text, XlsxDifferentialFontValue Font)> richRuns, List<PhoneticRunDto> phoneticRuns, PhoneticProperties? phoneticProps)
    {
        if (richRuns.Count == 0 && phoneticProps is null)
            return fullText ?? string.Empty;

        if (!string.IsNullOrEmpty(fullText))
            richRuns.Insert(0, (fullText, XlsxDifferentialFontValue.Empty));

        return GetRichText(richRuns, phoneticRuns, phoneticProps);
    }

    private static XlsxImmutableRichText GetRichText(List<(string Text, XlsxDifferentialFontValue Font)> richRuns, List<PhoneticRunDto> phoneticRunDtos, PhoneticProperties? phoneticProps)
    {
        var richTextLength = richRuns.Sum(x => x.Text.Length);
        var sb = new StringBuilder(richTextLength);
        foreach (var richRun in richRuns)
            sb.Append(richRun.Text);

        var richText = sb.ToString();

        // Phonetic runs must be in ascending order, non-overlapping
        // Any two consecutive phonetic runs should satisfy sb1 < eb1 <= sb2 < eb2 to avoid overlapping
        phoneticRunDtos.Sort(ComparePhoneticRunsByStartIndex);
        var prevStartIndex = richTextLength;
        for (var i = phoneticRunDtos.Count - 1; i >= 0; --i)
        {
            var phoneticRun = phoneticRunDtos[i];

            // Each run must have a text
            if (string.IsNullOrEmpty(phoneticRun.Text))
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            // Start index must be less than the end index
            if (phoneticRun.EndIndex < phoneticRun.StartIndex)
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            // Omit phonetic runs of length 0
            if (phoneticRun.EndIndex == phoneticRun.StartIndex)
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            // Omit phonetic runs that overlap with the next one
            if (phoneticRun.EndIndex > prevStartIndex)
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            prevStartIndex = phoneticRun.StartIndex;
        }

        var runs = new RichTextRun[richRuns.Count];
        var startIndex = 0;
        for (var i = 0; i < richRuns.Count; ++i)
        {
            var runLength = richRuns[i].Text.Length;
            runs[i] = new RichTextRun(XlsxFontFormatValue.Default, richRuns[i].Font, startIndex, runLength);
            startIndex += runLength;
        }

        var phoneticRuns = phoneticRunDtos.Select(x => new PhoneticRun(x.Text, x.StartIndex, x.EndIndex)).ToArray();
        return new XlsxImmutableRichText(richText, runs, phoneticRuns, phoneticProps);
    }
}
