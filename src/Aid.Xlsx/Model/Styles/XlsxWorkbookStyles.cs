using System.Diagnostics;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// A container for styles and formatting records in a workbook.
/// </summary>
internal class XlsxWorkbookStyles
{
    /// <summary>
    /// First user-defined numFmtId.
    /// </summary>
    public const int FirstUserDefinedNumberFormatIndex = 164;

    private readonly BiDictionary<int, XlsxNumberFormat> _numberFormats;

    private readonly BiDictionary<int, XlsxFontFormatValue> _fontFormats;

    private readonly BiDictionary<int, XlsxFillFormatValue> _fillFormats;

    private readonly BiDictionary<int, XlsxBorderFormatValue> _borderFormats;

    private readonly BiDictionary<int, XlsxAlignmentFormatValue> _alignmentFormats;

    private readonly BiDictionary<int, XlsxProtectionFormatValue> _protectionFormats;

    /// <summary>
    /// The key is XfId, the value is cell format.
    /// </summary>
    private readonly BiDictionary<int, XlsxCellFormatValue> _cellFormats;

    /// <summary>
    /// The key is cellStyleXfId, the value is cell style.
    /// </summary>
    private readonly BiDictionary<StyleId, XlsxCellStyleValue> _cellStyles;

    private readonly BiDictionary<int, XlsxDxfValue> _differentialFormats;

    /// <summary>
    /// Key is a table style name, value is a table style.
    /// </summary>
    private readonly Dictionary<string, XlsxTableTheme> _tableStyles;

    /// <summary>
    /// Key is a pivot table style name, value is a pivot table style.
    /// </summary>
    private readonly Dictionary<string, XlsxPivotTableStyle> _pivotStyles;

    private List<uint>? _indexedColorsArgb;

    private List<XlsxColor> _mruColors = new();

    /// <summary>
    /// A normal style that is used for newly create workbooks or loaded workbooks without a normal style.
    /// </summary>
    internal readonly XlsxCellStyleValue DefaultNormalStyle = new()
    {
        Name = "Normal",
        BuiltInStyle = BuiltInStyleValues.Normal,
        Hidden = false,
        Alignment = new XlsxAlignmentFormatValue
        {
            Horizontal = XlsxAlignmentHorizontalValues.General,
            Vertical = XlsxAlignmentVerticalValues.Bottom,
            TextRotation = TextRotation.None,
            WrapText = false,
            Indent = 0,
            RelativeIndent = 0,
            JustifyLastLine = false,
            ShrinkToFit = false,
            ReadingOrder = XlsxAlignmentReadingOrderValues.ContextDependent
        },
        Protection = new XlsxProtectionFormatValue
        {
            Locked = true,
            Hidden = false,
        },
        NumberFormat = XlsxPredefinedFormat.FormatCodes[XlsxPredefinedFormat.General],
        Font = new XlsxFontFormatValue
        {
            Name = "Calibri",
            Charset = XlsxFontCharSet.Ansi,
            Family = XlsxFontFamilyNumberingValues.Swiss,
            Bold = false,
            Italic = false,
            Strikethrough = false,
            Outline = false,
            Shadow = false,
            Condense = false,
            Extend = false,
            Color = XlsxColor.Black,
            Size = XlsxFontSize.FromPoints(11),
            Underline = XlsxFontUnderlineValues.None,
            VerticalAlignment = XlsxFontVerticalTextAlignmentValues.Baseline,
            Scheme = XlsxFontScheme.None
        },
        Fill = XlsxFillFormatValue.None,
        Border = XlsxBorderFormatValue.None,
        IncludedComponents = CellFormatComponents.All
    };

    internal XlsxWorkbookStyles()
    {
        _numberFormats = new BiDictionary<int, XlsxNumberFormat>();
        _fontFormats = new BiDictionary<int, XlsxFontFormatValue>();
        _fillFormats = new BiDictionary<int, XlsxFillFormatValue>();
        _borderFormats = new BiDictionary<int, XlsxBorderFormatValue>();
        _alignmentFormats = new BiDictionary<int, XlsxAlignmentFormatValue>();
        _protectionFormats = new BiDictionary<int, XlsxProtectionFormatValue>();
        _cellFormats = new BiDictionary<int, XlsxCellFormatValue>();
        _cellStyles = new BiDictionary<StyleId, XlsxCellStyleValue>();
        _differentialFormats = new BiDictionary<int, XlsxDxfValue>();
        _tableStyles = new Dictionary<string, XlsxTableTheme>(XlsxHelper.NameComparer);
        _pivotStyles = new Dictionary<string, XlsxPivotTableStyle>(XlsxHelper.NameComparer);
    }

    internal IReadOnlyBiDictionary<int, XlsxNumberFormat> NumberFormats => _numberFormats;

    internal IReadOnlyBiDictionary<int, XlsxFontFormatValue> Fonts => _fontFormats;

    internal IReadOnlyBiDictionary<int, XlsxFillFormatValue> Fills => _fillFormats;

    internal IReadOnlyBiDictionary<int, XlsxBorderFormatValue> Borders => _borderFormats;

    internal IReadOnlyBiDictionary<int, XlsxCellFormatValue> CellFormats => _cellFormats;

    internal IReadOnlyBiDictionary<StyleId, XlsxCellStyleValue> CellStyles => _cellStyles;

    internal IReadOnlyBiDictionary<int, XlsxDxfValue> DifferentialFormats => _differentialFormats;

    internal IReadOnlyDictionary<string, XlsxTableTheme> TableStyles => _tableStyles;

    internal IReadOnlyDictionary<string, XlsxPivotTableStyle> PivotStyles => _pivotStyles;

    /// <summary>
    /// Name of a table style that should be used for newly added tables. It's not used for tables
    /// without a specified style.
    /// </summary>
    internal string? DefaultTableStyle { get; set; }

    /// <summary>
    /// Name of a pivot style that should be used for newly added pivot tables. It's not used for
    /// pivot tables without a specified style.
    /// </summary>
    internal string? DefaultPivotStyle { get; set; }

    /// <summary>
    /// Some workbooks use indexed colors that are not in the standard <see cref="XlsxColor.IndexedColors"/>,
    /// but have their own list of indexed colors. Legacy feature, do not expose. If the value is null, use
    /// predefined indexed colors.
    /// </summary>
    internal IReadOnlyList<uint>? IndexedColorsArgb => _indexedColorsArgb;

    /// <summary>
    /// Most recently used colors are colors <em>hand-picked</em> by a user that are displayed in
    /// the color picker dialogue. The standard and theme colors are always offered in the color
    /// picker, so they are (in general) not added to the MRU color list.
    /// </summary>
    internal IReadOnlyList<XlsxColor> MruColors => _mruColors;

    /// <summary>
    /// The format inherited in the absence of an explicit format. It is initialized from cell
    /// format index zero on load and can be changed through <see cref="XlsxWorkbook.Style"/>.
    /// The current default is written at index zero on save.
    /// </summary>
    internal XlsxCellFormatValue DefaultFormat { get; set; } = new()
    {
        Font = new XlsxFontFormatValue
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
            Color = XlsxColor.FromArgb(0x00000000),
            Size = XlsxFontSize.FromPoints(11),
            Underline = XlsxFontUnderlineValues.None,
            VerticalAlignment = XlsxFontVerticalTextAlignmentValues.Baseline,
            Scheme = XlsxFontScheme.None
        },
        NumberFormat = XlsxPredefinedFormat.FormatCodes[XlsxPredefinedFormat.General],
        Alignment = new XlsxAlignmentFormatValue()
        {
            Horizontal = XlsxAlignmentHorizontalValues.General,
            Vertical = XlsxAlignmentVerticalValues.Bottom,
            TextRotation = TextRotation.None,
            WrapText = false,
            Indent = 0,
            RelativeIndent = 0,
            JustifyLastLine = false,
            ShrinkToFit = false,
            ReadingOrder = XlsxAlignmentReadingOrderValues.ContextDependent
        },
        Protection = new XlsxProtectionFormatValue
        {
            Locked = true,
            Hidden = false,
        },
        Fill = XlsxFillFormatValue.None,
        Border = XlsxBorderFormatValue.None,
        CellStyleId = null,
        IncludeQuotePrefix = false,
        PivotButton = false,
        CustomFormat = CellFormatComponents.None
    };

    internal void AddNumberFormat(int numFmtId, XlsxNumberFormat format)
    {
        _numberFormats.Add(numFmtId, format);
    }

    internal void AddUserDefinedNumberFormat(XlsxNumberFormat numberFormat)
    {
        var numFmtId = FirstUserDefinedNumberFormatIndex;
        if (_numberFormats.Count > 0)
            numFmtId = Math.Max(_numberFormats.Keys.Max() + 1, numFmtId);

        _numberFormats.Add(numFmtId, numberFormat);
    }

    internal void AddFontFormat(XlsxFontFormatValue fontFormat)
    {
        _fontFormats.Add(_fontFormats.Count, fontFormat);
    }

    internal void AddFillFormat(XlsxFillFormatValue fillFormat)
    {
        _fillFormats.Add(_fillFormats.Count, fillFormat);
    }

    internal void AddBorderFormat(XlsxBorderFormatValue borderFormat)
    {
        _borderFormats.Add(_borderFormats.Count, borderFormat);
    }

    internal void AddFormat(XlsxCellFormatValue cellFormat)
    {
        var xfId = _cellFormats.Count;
        _cellFormats.Add(xfId, cellFormat);
    }

    internal void AddCellStyle(int cellStyleXfId, XlsxCellStyleValue cellStyle)
    {
        _cellStyles.Add(cellStyleXfId, cellStyle);
    }

    internal void AddDifferentialFormat(XlsxDxfValue dxf)
    {
        _differentialFormats.Add(_differentialFormats.Count, dxf);
    }

    internal void AddTableStyle(XlsxTableTheme tableStyle)
    {
        _tableStyles.Add(tableStyle.Name, tableStyle);
    }

    internal void AddPivotStyle(XlsxPivotTableStyle pivotStyle)
    {
        _pivotStyles.Add(pivotStyle.Name, pivotStyle);
    }

    internal void SetIndexedColors(List<uint> indexedColors)
    {
        _indexedColorsArgb = indexedColors;
    }

    internal void SetMruColors(List<XlsxColor> mruColors)
    {
        _mruColors = mruColors;
    }

    internal XlsxNumberFormat RegisterNumberFormat(XlsxNumberFormat numberFormat)
    {
        if (_numberFormats.TryGetValue(numberFormat, out var existingFormat))
            return existingFormat;

        AddUserDefinedNumberFormat(numberFormat);
        return numberFormat;
    }

    private XlsxAlignmentFormatValue GetRegisteredAlignmentFormat(XlsxAlignmentFormatValue original, Func<XlsxAlignmentFormatValue, XlsxAlignmentFormatValue> modify)
    {
        return RegisterAlignmentFormat(modify(original));
    }

    internal XlsxAlignmentFormatValue RegisterAlignmentFormat(XlsxAlignmentFormatValue alignment)
    {
        if (_alignmentFormats.TryGetValue(alignment, out var existingAlignment))
            return existingAlignment;

        _alignmentFormats.Add(_alignmentFormats.Count, alignment);
        return alignment;
    }

    internal XlsxProtectionFormatValue RegisterProtectionFormat(XlsxProtectionFormatValue protection)
    {
        if (_protectionFormats.TryGetValue(protection, out var existingProtection))
            return existingProtection;

        _protectionFormats.Add(_protectionFormats.Count, protection);
        return protection;
    }

    /// <summary>
    /// Get a font format that is stored in the internal structures of the styles class. The font
    /// format is created by modification of existing font format. This is essential for saving,
    /// all formats must be registered in the styles class. 
    /// </summary>
    internal XlsxFontFormatValue GetRegisteredFontFormat(XlsxFontFormatValue original, Func<XlsxFontFormatValue, XlsxFontFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterFontFormat(modified);
    }

    internal XlsxFontFormatValue RegisterFontFormat(XlsxFontFormatValue font)
    {
        if (_fontFormats.TryGetValue(font, out var existingFont))
            return existingFont;

        AddFontFormat(font);
        return font;
    }

    internal XlsxCellFormatValue GetModifiedFormat(XlsxCellFormatValue originalFormat, XlsxNumberFormat numberFormat)
    {
        var modifiedNumberFormat = RegisterNumberFormat(numberFormat);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { NumberFormat = modifiedNumberFormat });
        return modifiedFormat;
    }

    internal XlsxCellFormatValue GetModifiedFormat(XlsxCellFormatValue originalFormat, Func<XlsxAlignmentFormatValue, XlsxAlignmentFormatValue> modify)
    {
        var modifiedAlignment = GetRegisteredAlignmentFormat(originalFormat.Alignment, modify);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { Alignment = modifiedAlignment });
        return modifiedFormat;
    }

    internal XlsxCellFormatValue GetModifiedFormat(XlsxCellFormatValue originalFormat, Func<XlsxFontFormatValue, XlsxFontFormatValue> modify)
    {
        var modifiedFont = GetRegisteredFontFormat(originalFormat.Font, modify);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { Font = modifiedFont });
        return modifiedFormat;
    }

    internal XlsxCellFormatValue GetModifiedFormat(XlsxCellFormatValue originalFormat, Func<XlsxBorderFormatValue, XlsxBorderFormatValue> modify)
    {
        var modifiedBorder = GetRegisteredBorderFormat(originalFormat.Border, modify);
        var modifiedFormat = GetRegisteredCellFormat(originalFormat, format => format with { Border = modifiedBorder });
        return modifiedFormat;
    }

    internal XlsxFillFormatValue GetRegisteredFillFormat(XlsxFillFormatValue original, Func<XlsxFillFormatValue, XlsxFillFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterFillFormat(modified);
    }

    private XlsxFillFormatValue RegisterFillFormat(XlsxFillFormatValue fill)
    {
        if (_fillFormats.TryGetValue(fill, out var existingFill))
            return existingFill;

        AddFillFormat(fill);
        return fill;
    }

    internal XlsxBorderFormatValue GetRegisteredBorderFormat(XlsxBorderFormatValue original, Func<XlsxBorderFormatValue, XlsxBorderFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterBorderFormat(modified);
    }

    private XlsxBorderFormatValue RegisterBorderFormat(XlsxBorderFormatValue border)
    {
        if (_borderFormats.TryGetValue(border, out var existingBorder))
            return existingBorder;

        AddBorderFormat(border);
        return border;
    }

    internal XlsxCellFormatValue GetRegisteredCellFormat(XlsxCellFormatValue original, Func<XlsxCellFormatValue, XlsxCellFormatValue> modify)
    {
        var modified = modify(original);
        return RegisterCellFormat(modified);
    }

    internal XlsxCellFormatValue RegisterCellFormat(XlsxCellFormatValue cellFormat)
    {
        if (_cellFormats.TryGetValue(cellFormat, out var existing))
            return existing;

        Debug.Assert(_numberFormats.ContainsValue(cellFormat.NumberFormat));
        Debug.Assert(_alignmentFormats.TryGetValue(cellFormat.Alignment, out var registeredAlignment) && ReferenceEquals(cellFormat.Alignment, registeredAlignment));
        Debug.Assert(_protectionFormats.TryGetValue(cellFormat.Protection, out var registeredProtection) && ReferenceEquals(cellFormat.Protection, registeredProtection));
        Debug.Assert(_fontFormats.TryGetValue(cellFormat.Font, out var registeredFont) && ReferenceEquals(cellFormat.Font, registeredFont));
        Debug.Assert(_fillFormats.TryGetValue(cellFormat.Fill, out var registeredFill) && ReferenceEquals(cellFormat.Fill, registeredFill));
        Debug.Assert(_borderFormats.TryGetValue(cellFormat.Border, out var registeredBorder) && ReferenceEquals(cellFormat.Border, registeredBorder));
        AddFormat(cellFormat);
        return cellFormat;
    }

    /// <summary>
    /// Get registered format equal to <paramref name="format"/> from the styles. Generally for copying formats other workbooks.
    /// </summary>
    internal XlsxCellFormatValue GetRegisteredCellFormat(XlsxCellFormatValue format)
    {
        // If format is already registered, all its components must be registered too.
        if (_cellFormats.TryGetValue(format, out var existing))
            return existing;

        // TODO: If format is from different workbook, we should copy style if from different workbook and this workbook doesn't already have a style with same name
        StyleId? formatStyleId = format.CellStyleId is { } cellStyleId && _cellStyles.ContainsKey(cellStyleId) ? cellStyleId : null;

        // We have to create new one, because some components may already exist here
        var cellFormat = new XlsxCellFormatValue
        {
            NumberFormat = RegisterNumberFormat(format.NumberFormat),
            Alignment = RegisterAlignmentFormat(format.Alignment),
            Protection = RegisterProtectionFormat(format.Protection),
            Font = RegisterFontFormat(format.Font),
            Fill = RegisterFillFormat(format.Fill),
            Border = RegisterBorderFormat(format.Border),
            CellStyleId = formatStyleId,
            IncludeQuotePrefix = format.IncludeQuotePrefix,
            PivotButton = format.PivotButton,
            CustomFormat = format.CustomFormat
        };

        AddFormat(cellFormat);
        return cellFormat;
    }

    /// <summary>
    /// Get a differential format that is stored in the internal structures of the styles class.
    /// The differential format is created by modification of existing dxf format. This is
    /// essential for saving, all formats must be registered in the styles class. 
    /// </summary>
    internal XlsxDxfValue GetRegisteredDxFormat(XlsxDxfValue original, Func<XlsxDxfValue, XlsxDxfValue> modify)
    {
        var modified = modify(original);
        return RegisterDxFormat(modified);
    }

    /// <summary>
    /// Register dxf from potentially different workbook into this workbook.
    /// </summary>
    /// <returns>Registered instance.</returns>
    internal XlsxDxfValue RegisterDxFormat(XlsxDxfValue dxf)
    {
        if (_differentialFormats.TryGetValue(dxf, out var existingDxf))
            return existingDxf;

        AddDifferentialFormat(dxf);
        return dxf;
    }

    /// <summary>
    /// Initialize styles component so it does contain styles suitable for a new workbook.
    /// </summary>
    internal void Initialize()
    {
        DefaultTableStyle = XlsxTableTheme.TableStyleMedium2.ToString();
        DefaultPivotStyle = nameof(XlsxPivotTableTheme.PivotStyleLight16);

        foreach (var (numFmtId, formatCode) in XlsxPredefinedFormat.FormatCodes)
            AddNumberFormat(numFmtId, formatCode);

        var normalStyle = DefaultNormalStyle;
        AddFontFormat(normalStyle.Font);
        AddFillFormat(XlsxFillFormatValue.None);
        AddFillFormat(XlsxFillFormatValue.Gray125);
        AddBorderFormat(XlsxBorderFormatValue.None);
        AddCellStyle(0, normalStyle);

        var defaultFormat = XlsxCellFormatValue.FromStyle(0, normalStyle);
        DefaultFormat = GetRegisteredCellFormat(defaultFormat);
    }
}
