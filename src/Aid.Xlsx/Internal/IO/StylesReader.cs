using System.Diagnostics;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Opc;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles.Formatting;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Tables;
using Aid.Xlsx.Utils;
using PTS = Aid.Xlsx.Styles.Formatting.XlsxPivotStyleRegionValues;
using TS = Aid.Xlsx.Styles.Formatting.XlsxTableStyleRegionValues;

namespace Aid.Xlsx.Internal.IO;

internal partial class StylesReader
{
    private readonly XmlTreeReader _reader;
    private readonly XlsxWorkbookStyles _styles;
    private readonly string _ns = OpenXmlConst.Main2006SsNs;
    private readonly SequentialNameGenerator _styleNameGenerator = new("Style ", 1);

    // Format components to use when not specified in xf record
    private readonly XlsxNumberFormat _defaultNumberFormat;
    private readonly XlsxFillFormatValue _defaultFillFormat;
    private readonly XlsxBorderFormatValue _defaultBorderFormat;
    private readonly XlsxAlignmentFormatValue _defaultAlignmentFormat;
    private readonly XlsxProtectionFormatValue _defaultProtectionFormat;
    private XlsxFontFormatValue _defaultFontFormat;

    // Currently read CT_TableStyle element
    private Dictionary<TS, (XlsxDxfValue Dxf, int BandSize)> _currentTableStyle = new();
    private Dictionary<PTS, (XlsxDxfValue Dxf, int BandSize)> _currentPivotStyle = new();

    /// <summary>
    /// Style formats from <c>cellStyleXfs</c>.
    /// </summary>
    private List<XlsxCellFormatValue> _styleFormats = new();

    public StylesReader(XmlTreeReader reader, XlsxWorkbookStyles styles)
    {
        _reader = reader;
        _styles = styles;

        // Set initial fallback values if part is empty
        _defaultNumberFormat = styles.DefaultNormalStyle.NumberFormat;
        _defaultFontFormat = styles.DefaultNormalStyle.Font;
        _defaultFillFormat = styles.DefaultNormalStyle.Fill;
        _defaultBorderFormat = styles.DefaultNormalStyle.Border;
        _defaultAlignmentFormat = styles.DefaultNormalStyle.Alignment;
        _defaultProtectionFormat = styles.DefaultNormalStyle.Protection;
    }

    internal void Load()
    {
        _reader.Open("styleSheet", _ns);
        ParseStylesheet("styleSheet");

        LoadDefaultFormat();
    }

    private void LoadDefaultFormat()
    {
        // Normal style is technically optional, but it's basically a requirement for any sane work.
        var (styleId, normalStyle) = _styles.CellStyles.SingleOrDefault(x => x.Value.BuiltInStyle == BuiltInStyleValues.Normal);
        if (normalStyle is null)
        {
            styleId = _styles.CellStyles.Count;
            normalStyle = _styles.DefaultNormalStyle;

            // Number format collection already contains all predefined numFmts, so it can only be user-defined
            if (!_styles.NumberFormats.ContainsValue(normalStyle.NumberFormat))
                _styles.AddUserDefinedNumberFormat(normalStyle.NumberFormat);

            if (!_styles.Fonts.ContainsValue(normalStyle.Font))
                _styles.AddFontFormat(normalStyle.Font);

            if (!_styles.Fills.ContainsValue(normalStyle.Fill))
                _styles.AddFillFormat(normalStyle.Fill);

            if (!_styles.Borders.ContainsValue(normalStyle.Border))
                _styles.AddBorderFormat(normalStyle.Border);

            _styles.AddCellStyle(styleId.Value, normalStyle);
        }

        // Ensure there is a default format.
        if (!_styles.CellFormats.TryGetValue(0, out var defaultFormat))
        {
            defaultFormat = XlsxCellFormatValue.FromStyle(styleId, normalStyle);
            _styles.AddFormat(defaultFormat);
        }

        _styles.DefaultFormat = defaultFormat;
    }

    private void ParseStylesheet(string elementName)
    {
        ParseNumFmts("numFmts", _ns);

        // The spec says that the predefined formats have "formatCode value [..] implied rather
        // than explicitly saved in the file."... so if there was something saved, it should have
        // be used. If something was saved, it's explicit and explicit (generally) has preference
        // over implicit. It needs to be added after numFmts, but before cellStyleXfs/cellXfs.
        AddImpliedNumberFormats();

        ParseFonts("fonts", _ns);

        if (_styles.Fonts.Count == 0)
        {
            _styles.AddFontFormat(_defaultFontFormat);
        }
        else
        {
            _defaultFontFormat = _styles.Fonts[0];
        }

        ParseFills("fills", _ns);

        // Default fill is always none, should be at index 0.
        if (!_styles.Fills.ContainsValue(_defaultFillFormat))
        {
            _styles.AddFillFormat(_defaultFillFormat);
        }

        ParseBorders("borders", _ns);

        if (!_styles.Borders.ContainsValue(_defaultBorderFormat))
        {
            _styles.AddBorderFormat(_defaultBorderFormat);
        }

        ParseCellStyleXfs("cellStyleXfs", _ns);

        var cellFormats = new List<(XlsxCellFormatValue Format, int? CellStyleXfId)>();
        if (ParseCellXfs("cellXfs", _ns) is { IsSuccess: true } cellXfsResult)
        {
            cellFormats = cellXfsResult.Value;
        }

        var cellStyles = new Dictionary<int, XlsxCellStyleValue>();
        if (ParseCellStyles("cellStyles", _ns) is { IsSuccess: true } cellStylesResult)
        {
            cellStyles = cellStylesResult.Value;
        }

        RepairMissingStyles(cellStyles);
        AddCellStyles(cellStyles);
        AddFormats(cellFormats, cellStyles);

        ParseDxfs("dxfs", _ns);
        ParseTableStyles("tableStyles", _ns);
        ParseColors("colors", _ns);
        ParseExtensionList("extLst", _ns);
        _reader.Close(elementName, _ns);
    }

    private void AddImpliedNumberFormats()
    {
        // Add predefined formats, so we can treat predefined number formats and
        // user defined number formats the same way.
        foreach (var (numFmtId, formatCode) in XlsxPredefinedFormat.FormatCodes)
        {
            if (!_styles.NumberFormats.ContainsKey(numFmtId))
            {
                _styles.AddNumberFormat(numFmtId, formatCode);
            }
        }
    }

    private void RepairMissingStyles(Dictionary<int, XlsxCellStyleValue> cellStyles)
    {
        // Because cellStyleXfs might be referenced from cell formats, each one must be converted
        // to a cell style. If the cellStyles didn't contain a record for any cellStyleXf, add it.
        for (var cellStyleXfId = 0; cellStyleXfId < _styleFormats.Count; ++cellStyleXfId)
        {
            if (!cellStyles.ContainsKey(cellStyleXfId))
            {
                var format = _styleFormats[cellStyleXfId];
                var generatedName = _styleNameGenerator.NextUnusedStyleName();
                cellStyles.Add(cellStyleXfId, new XlsxCellStyleValue
                {
                    Name = generatedName,
                    BuiltInStyle = null,
                    Hidden = false,
                    NumberFormat = format.NumberFormat,
                    Alignment = format.Alignment,
                    Protection = format.Protection,
                    Font = format.Font,
                    Fill = format.Fill,
                    Border = format.Border,
                    IncludedComponents = CellFormatComponents.All
                });
            }
        }
    }

    private void AddCellStyles(Dictionary<int, XlsxCellStyleValue> cellStyles)
    {
        foreach (var (cellStyleXfId, cellStyle) in cellStyles)
            _styles.AddCellStyle(cellStyleXfId, cellStyle);
    }

    private void AddFormats(List<(XlsxCellFormatValue Format, int? CellStyleXfId)> cellFormats, Dictionary<int, XlsxCellStyleValue> cellStyles)
    {
        // At the time when cellXf were parsed, cell styles weren't resolved. Resolve them now.
        for (var xfId = 0; xfId < cellFormats.Count; ++xfId)
        {
            if (cellFormats[xfId].CellStyleXfId is { } cellStyleXfId)
            {
                // Make sure a the styleId is valid. The styleId for cell formats is read from file
                // and thus could be invalid. We don't want to crash later.
                if (!cellStyles.ContainsKey(cellStyleXfId))
                    throw PartStructureException.InvalidAttributeValue();

                var cellFormat = cellFormats[xfId].Format with { CellStyleId = cellStyleXfId };
                cellFormats[xfId] = (cellFormat, cellStyleXfId);
            }
        }

        foreach (var (cellFormat, _) in cellFormats)
            _styles.AddFormat(cellFormat);
    }

    private (int NumFmtId, XlsxNumberFormat FormatCode) OnNumFmtParsed(uint numFmtId, string formatCode)
    {
        return (checked((int)numFmtId), XlsxNumberFormat.Parse(formatCode));
    }

    partial void OnNumFmtsParsed(List<(int NumFmtId, XlsxNumberFormat Format)> numFmt, uint? count)
    {
        foreach (var (numFmtId, formatCode) in numFmt)
        {
            // Even if numFmtId is predefined, store the supplied value. Excel accepts predefined
            // numFmtId and uses supplied format instead of predefined format. It fixes situation
            // during save, where such items are saved to user supplied numFmtId range.
            _styles.AddNumberFormat(numFmtId, formatCode);
        }
    }

    private Xpr<XlsxDifferentialFontValue> ParseFont(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxDifferentialFontValue>();
        }

        // Font is mostly buggy specification. Excel basically chokes on anything but a sequence,
        // but standard requires an unbound choice where elements can repeat.
        XlsxFontName? fontName = null;
        XlsxFontCharSet? fontCharset = null;
        XlsxFontFamilyNumberingValues? fontFamily = null;
        bool? fontBold = null, fontItalic = null, fontStrikethrough = null, fontOutline = null, fontShadow = null, fontCondense = null, fontExtend = null;
        XlsxColor? fontColor = null;
        XlsxFontSize? fontSize = null;
        XlsxFontUnderlineValues? fontUnderline = null;
        XlsxFontVerticalTextAlignmentValues? fontVerticalAlignment = null;
        XlsxFontScheme? fontScheme = null;
        while (!_reader.TryClose(elementName, ns))
        {
            if (_reader.TryReadXStringValElement("name", _ns, out var name))
            {
                fontName = name;
            }
            else if (_reader.TryReadIntValElement("charset", _ns, out var charset))
            {
                fontCharset = (XlsxFontCharSet)charset;
            }
            else if (_reader.TryReadIntValElement("family", _ns, out var family))
            {
                // Bug in the spec. Spec says that it has values 0-14 and doesn't specify meaning
                // for the numerical values. It's supposed to refer to the same enum ST_FontFamily
                // as in WordML. The OI-29500 fixes this problem:
                // "Excel restricts the value of this attribute to be at least 0 and at most 5."
                fontFamily = family switch
                {
                    >= 0 and <= 5 => (XlsxFontFamilyNumberingValues)family,
                    > 5 and <= 14 => XlsxFontFamilyNumberingValues.NotApplicable,
                    _ => throw PartStructureException.InvalidAttributeFormat(),
                };
            }
            else if (_reader.TryReadBoolElement("b", _ns, out var b))
            {
                fontBold = b;
            }
            else if (_reader.TryReadBoolElement("i", _ns, out var i))
            {
                fontItalic = i;
            }
            else if (_reader.TryReadBoolElement("strike", _ns, out var strike))
            {
                fontStrikethrough = strike;
            }
            else if (_reader.TryReadBoolElement("outline", _ns, out var outline))
            {
                fontOutline = outline;
            }
            else if (_reader.TryReadBoolElement("shadow", _ns, out var shadow))
            {
                fontShadow = shadow;
            }
            else if (_reader.TryReadBoolElement("condense", _ns, out var condense))
            {
                fontCondense = condense;
            }
            else if (_reader.TryReadBoolElement("extend", _ns, out var extend))
            {
                fontExtend = extend;
            }
            else if (_reader.TryReadColor("color", _ns, out var color))
            {
                fontColor = color;
            }
            else if (_reader.TryOpen("sz", _ns))
            {
                var fontSizePt = _reader.GetDouble("val");
                _reader.Close("sz", _ns);
                fontSize = XlsxFontSize.FromPoints(fontSizePt);
            }
            else if (_reader.TryOpen("u", _ns))
            {
                var underline = _reader.GetOptionalEnum<XlsxFontUnderlineValues>("val") ?? XlsxFontUnderlineValues.Single;
                _reader.Close("u", _ns);
                fontUnderline = underline;
            }
            else if (_reader.TryReadEnumValElement<XlsxFontVerticalTextAlignmentValues>("vertAlign", _ns, out var vertAlign))
            {
                fontVerticalAlignment = vertAlign;
            }
            else if (_reader.TryReadEnumValElement<XlsxFontScheme>("scheme", _ns, out var scheme))
            {
                fontScheme = scheme;
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(_reader);
            }
        }

        var fontFormat = new XlsxDifferentialFontValue
        {
            Name = fontName,
            Charset = fontCharset,
            Family = fontFamily,
            Bold = fontBold,
            Italic = fontItalic,
            Strikethrough = fontStrikethrough,
            Outline = fontOutline,
            Shadow = fontShadow,
            Condense = fontCondense,
            Extend = fontExtend,
            Color = fontColor,
            Size = fontSize,
            Underline = fontUnderline,
            VerticalAlignment = fontVerticalAlignment,
            Scheme = fontScheme,
        };
        return Xpr.From(fontFormat);
    }

    // ParseFont is shared between <fonts> table and <dxf> elements. Once the <fonts> table is read,
    // register collected fonts to the workbook styles.
    partial void OnFontsParsed(List<XlsxDifferentialFontValue> font, uint? count)
    {
        var defaultFont = XlsxFontFormatValue.Default;

        // Excel probably screwed up. Default name and size should likely be font of default
        // format, but it is always taken from font zero (if font zero defines name/size).
        if (font.Count > 0)
        {
            var fontZero = font[0];
            defaultFont = defaultFont with
            {
                Name = fontZero.Name ?? defaultFont.Name,
                Size = fontZero.Size ?? defaultFont.Size,
            };
        }

        foreach (var fontProps in font)
        {
            var fontFormat = new XlsxFontFormatValue
            {
                Name = fontProps.Name ?? defaultFont.Name,
                Size = fontProps.Size ?? defaultFont.Size,
                Charset = fontProps.Charset ?? defaultFont.Charset,
                Family = fontProps.Family ?? defaultFont.Family,
                Bold = fontProps.Bold ?? defaultFont.Bold,
                Italic = fontProps.Italic ?? defaultFont.Italic,
                Strikethrough = fontProps.Strikethrough ?? defaultFont.Strikethrough,
                Outline = fontProps.Outline ?? defaultFont.Outline,
                Shadow = fontProps.Shadow ?? defaultFont.Shadow,
                Condense = fontProps.Condense ?? defaultFont.Condense,
                Extend = fontProps.Extend ?? defaultFont.Extend,
                Color = fontProps.Color ?? defaultFont.Color,
                Underline = fontProps.Underline ?? defaultFont.Underline,
                VerticalAlignment = fontProps.VerticalAlignment ?? defaultFont.VerticalAlignment,
                Scheme = fontProps.Scheme ?? defaultFont.Scheme,
            };
            _styles.AddFontFormat(fontFormat);
        }
    }

    private XlsxFillFormatValue OnFillParsed(XlsxFillFormatValue? foundFill)
    {
        var fillFormat = foundFill ?? XlsxFillFormatValue.Empty;
        _styles.AddFillFormat(fillFormat);
        return fillFormat;
    }

    private XlsxFillFormatValue OnPatternFillParsed(XlsxColor? fgColor, XlsxColor? bgColor, XlsxFillPatternValues? patternType)
    {
        // There is a discrepancy between <fill> interpretation for a solid fill:
        // * cell fill: Pattern color is ignored, the background color is used for fill
        // * dxf fill: Pattern color is the one used for fill, the background color is ignored
        // The GUI in both cases says that the background color is the one that is used. Therefore
        // use background is correct per GUI. The problem is that the ported code historically says
        // the pattern color is the one that is used. This sucks, I have to live with it.
        // The alternative is a breaking change with no benefit.
        //
        // The other difference between dxf and cell fill is the default pattern type. Spec and
        // OI-29500 is silent, but Excel uses solid fill for dxf and none for cell fill.
        if (_reader.Context[^2] == "dxf")
        {
            var patternFill = new XlsxPatternFill
            {
                PatternColor = fgColor ?? XlsxColor.Automatic,
                BackgroundColor = bgColor ?? XlsxColor.Automatic,
                PatternType = patternType ?? XlsxFillPatternValues.Solid,
            };
            return new XlsxFillFormatValue(patternFill);
        }
        else
        {
            var pattern = patternType ?? XlsxFillPatternValues.None;

            // Fix solid pattern discrepancy for cell fill
            if (pattern == XlsxFillPatternValues.Solid)
                (bgColor, fgColor) = (fgColor, bgColor);

            var patternFill = new XlsxPatternFill
            {
                PatternColor = fgColor ?? XlsxColor.Automatic,
                BackgroundColor = bgColor ?? XlsxColor.Automatic,
                PatternType = pattern,
            };
            return new XlsxFillFormatValue(patternFill);
        }
    }

    private XlsxFillFormatValue OnGradientFillParsed(List<(FractionOfOne Value, XlsxColor Color)> stop, XlsxGradientType type, double degree, double left, double right, double top, double bottom)
    {
        var stops = stop.ToDictionary(x => x.Value, x => x.Color);
        switch (type)
        {
            case XlsxGradientType.Linear:
                return new XlsxFillFormatValue(new XlsxLinearGradientFill
                {
                    Stops = stops,
                    Degrees = degree,
                });
            case XlsxGradientType.Path:
                return new XlsxFillFormatValue(new XlsxPathGradientFill
                {
                    Stops = stops,
                    InnerLeft = left,
                    InnerRight = right,
                    InnerTop = top,
                    InnerBottom = bottom
                });
            default:
                throw new UnreachableException();
        }
    }

    private (FractionOfOne Position, XlsxColor Color) OnGradientStopParsed(XlsxColor color, double position)
    {
        // Spec requires stop positions to be 0..1, but doesn't have a type for that. Excel repairs workbook when it receives values outside 0..1.
        return (position, color);
    }

    private XlsxDifferentialBorderValue OnBorderParsed(XlsxBorderLine? start, XlsxBorderLine? end, XlsxBorderLine? left, XlsxBorderLine? right, XlsxBorderLine? top, XlsxBorderLine? bottom, XlsxBorderLine? diagonal, XlsxBorderLine? vertical, XlsxBorderLine? horizontal, bool? diagonalUp, bool? diagonalDown, bool outline)
    {
        var dxfBorder = new XlsxDifferentialBorderValue
        {
            Left = left ?? start,
            Right = right ?? end,
            Top = top,
            Bottom = bottom,
            Diagonal = diagonal,
            Vertical = vertical,
            Horizontal = horizontal,
            DiagonalUp = diagonalUp ?? false,
            DiagonalDown = diagonalDown ?? false,
            Outline = outline
        };
        if (_reader.Context[^1] == "borders")
        {
            var cellBorder = XlsxBorderFormatValue.FromDxf(dxfBorder);
            _styles.AddBorderFormat(cellBorder);
        }

        return dxfBorder;
    }

    private XlsxBorderLine OnBorderPrParsed(XlsxColor? color, XlsxBorderStyleValues style)
    {
        return new XlsxBorderLine(color ?? XlsxColor.Automatic, style);
    }

    partial void OnCellStyleXfsParsed(List<(XlsxCellFormatValue Format, int? CellStyleXfId)> xf, uint? count)
    {
        _styleFormats = xf.Select(x => x.Format).ToList();
    }

    private (XlsxCellFormatValue Format, int? CellStyleXfId) OnXfParsed(XlsxDifferentialAlignmentValue? alignment, XlsxDifferentialProtectionValue? protection, uint? numFmtId, uint? fontId, uint? fillId, uint? borderId, uint? xfId, bool quotePrefix, bool pivotButton, bool? applyNumberFormat, bool? applyFont, bool? applyFill, bool? applyBorder, bool? applyAlignment, bool? applyProtection)
    {
        // When xf is parsed, all number formats, fonts, fills and borders should already be read.
        var numberFormat = _defaultNumberFormat;
        if (numFmtId is not null && _styles.NumberFormats.TryGetValue(checked((int)numFmtId), out var numFmt))
            numberFormat = numFmt;

        var font = fontId is not null ? _styles.Fonts[checked((int)fontId)] : _defaultFontFormat;
        var fill = fillId is not null ? _styles.Fills[checked((int)fillId)] : _defaultFillFormat;
        var border = borderId is not null ? _styles.Borders[checked((int)borderId)] : _defaultBorderFormat;

        // Excel doesn't actually use the apply* for xf, but at least it writes as if it did. It
        // actually checks whether the id is same for xf and a style and if it is, the aspect
        // should be from a style. Excel is doesn't use it, other producers might.

        // Cell format has default apply* false (interpreted as "does format has its own custom format for this aspect")
        // Style has default apply* true (interpreted as "does style define this aspect")
        // The apply* attributes have default value true for cellStyleXfs and false for cellXfs.
        var isStyleXf = _reader.Context[^1] == "cellStyleXfs";
        var isCellFormat = !isStyleXf;
        var defaultApply = !isCellFormat;
        var components = CellFormatComponents.None;

        if (applyNumberFormat ?? defaultApply)
            components |= CellFormatComponents.NumberFormat;

        if (applyFont ?? defaultApply)
            components |= CellFormatComponents.Font;

        if (applyFill ?? defaultApply)
            components |= CellFormatComponents.Fill;

        if (applyBorder ?? defaultApply)
            components |= CellFormatComponents.Border;

        if (applyAlignment ?? defaultApply)
            components |= CellFormatComponents.Alignment;

        if (applyProtection ?? defaultApply)
            components |= CellFormatComponents.Protection;

        var formatAlignment = alignment is not null ? new XlsxAlignmentFormatValue
        {
            Horizontal = alignment.Horizontal ?? XlsxAlignmentFormatValue.Default.Horizontal,
            Vertical = alignment.Vertical ?? XlsxAlignmentFormatValue.Default.Vertical,
            TextRotation = alignment.TextRotation ?? XlsxAlignmentFormatValue.Default.TextRotation,
            WrapText = alignment.WrapText ?? XlsxAlignmentFormatValue.Default.WrapText,
            Indent = alignment.Indent ?? XlsxAlignmentFormatValue.Default.Indent,
            RelativeIndent = alignment.RelativeIndent ?? XlsxAlignmentFormatValue.Default.RelativeIndent,
            JustifyLastLine = alignment.JustifyLastLine ?? XlsxAlignmentFormatValue.Default.JustifyLastLine,
            ShrinkToFit = alignment.ShrinkToFit ?? XlsxAlignmentFormatValue.Default.ShrinkToFit,
            ReadingOrder = alignment.ReadingOrder ?? XlsxAlignmentFormatValue.Default.ReadingOrder,
        } : XlsxAlignmentFormatValue.Default;
        var formatProtection = protection is not null ? new XlsxProtectionFormatValue
        {
            Locked = protection.Locked ?? XlsxProtectionFormatValue.Default.Locked,
            Hidden = protection.Hidden ?? XlsxProtectionFormatValue.Default.Hidden
        } : XlsxProtectionFormatValue.Default;
        var format = new XlsxCellFormatValue
        {
            NumberFormat = numberFormat,
            // Alignment is not copied from default format
            Alignment = _styles.RegisterAlignmentFormat(formatAlignment),
            Protection = _styles.RegisterProtectionFormat(formatProtection),
            Font = font,
            Fill = fill,
            Border = border,
            CellStyleId = null, // The style is set once cell styles are resolved
            IncludeQuotePrefix = quotePrefix,
            PivotButton = pivotButton,
            CustomFormat = components
        };
        return (format, checked((int?)xfId));
    }

    private List<(XlsxCellFormatValue Format, int? CellStyleXfId)> OnCellXfsParsed(List<(XlsxCellFormatValue Format, int? CellStyleXfId)> xf, uint? count)
    {
        return xf;
    }

    private (int XfId, XlsxCellStyleValue Style) OnCellStyleParsed(string? name, uint xfId, uint? builtinId, uint? iLevel, bool? hidden, bool? customBuiltin)
    {
        // The quotePrefix and pivotButton attributes of the cellStyleXf are not applied, plus
        // the xfId of the cellStyle is ignored. The OI-29500 also requires uniqueness of xfId
        // in the cellStyle elements, although Excel can load such workbook and has several
        // "linked" styles. It's likely the separation into two elements is based on the internal
        // structure inside the Excel.

        // Fill dummy name for the style
        if (string.IsNullOrWhiteSpace(name))
        {
            name = _styleNameGenerator.NextUnusedStyleName();
        }
        else
        {
            _styleNameGenerator.AddName(name);
        }

        var cellStyleFormat = _styleFormats[checked((int)xfId)];

        // If the built in style is an outline style, expand it to avoid ugly representation.
        // The spec has only one only builtIn id for all RowLevel* styles (1) and one builtIn
        // id for all ColLevel* styles (2). Since the iLevel/outlineLevel is used only for
        // the RowLevel/ColLevel (OI-29500), expand the builtIn+iLevel for Row/Col level into
        // a separate builtIn styles (101-107 for RowLevel1-7, 201-207 for ColumnLevel1-7).
        if (builtinId is 1 or 2)
            builtinId = builtinId.Value * 100 + 1 + iLevel ?? 0;

        // BuiltIn must be among defined built-in styles ("implementers should restrict the content
        // of this attribute to enumerations present in the list")
        if (builtinId is not null && !Enum.IsDefined(typeof(BuiltInStyleValues), checked((int)builtinId.Value)))
        {
            throw PartStructureException.InvalidAttributeFormat();
        }

        // The apply* attributes have default `true` for cellStyleXfs and `false` for cellXfs.
        // We already took care of correct default value during the parsing of <xf>, so we don't
        // have to deal with it here.
        var styleIncludesComponents = cellStyleFormat.CustomFormat;
        var cellStyle = new XlsxCellStyleValue
        {
            Name = name,
            BuiltInStyle = builtinId is not null ? (BuiltInStyleValues)builtinId.Value : null,
            Hidden = hidden ?? false,
            NumberFormat = cellStyleFormat.NumberFormat,
            Alignment = cellStyleFormat.Alignment,
            Protection = cellStyleFormat.Protection,
            Font = cellStyleFormat.Font,
            Fill = cellStyleFormat.Fill,
            Border = cellStyleFormat.Border,
            IncludedComponents = styleIncludesComponents
        };

        return (checked((int)xfId), cellStyle);
    }

    private Dictionary<int, XlsxCellStyleValue> OnCellStylesParsed(List<(int CellStyleXfId, XlsxCellStyleValue Style)> cellStyle, uint? count)
    {
        var cellStyles = new Dictionary<int, XlsxCellStyleValue>();
        foreach (var (cellStyleXfId, style) in cellStyle)
        {
            // Multiple cell styles use same style formatting - split them, so each one uses
            // separate formatting. I considered removing duplicates, but it could mean that I
            // also might remove normal style, which is not desirable.
            if (!cellStyles.ContainsKey(cellStyleXfId))
            {
                cellStyles.Add(cellStyleXfId, style);
            }
            else
            {
                _styleFormats.Add(_styleFormats[cellStyleXfId]);
                var newCellStyleXfId = _styleFormats.Count - 1;
                cellStyles.Add(newCellStyleXfId, style);
            }
        }

        return cellStyles;
    }

    private XlsxDifferentialAlignmentValue OnCellAlignmentParsed(XlsxAlignmentHorizontalValues? horizontal, XlsxAlignmentVerticalValues vertical, uint? textRotation, bool? wrapText, uint? indent, int? relativeIndent, bool? justifyLastLine, bool? shrinkToFit, uint? readingOrder)
    {
        if (readingOrder is not null && readingOrder is not (0 or 1 or 2))
            throw PartStructureException.InvalidAttributeFormat();

        var normalizedTextRotation = OpenXmlHelper.NormalizeRotation(textRotation ?? 0);
        return new XlsxDifferentialAlignmentValue
        {
            Horizontal = horizontal,
            Vertical = vertical,
            TextRotation = textRotation is not null ? new TextRotation(normalizedTextRotation) : null,
            WrapText = wrapText,
            Indent = indent is not null ? checked((int)indent.Value) : null,
            RelativeIndent = relativeIndent,
            JustifyLastLine = justifyLastLine,
            ShrinkToFit = shrinkToFit,
            ReadingOrder = readingOrder is not null ? (XlsxAlignmentReadingOrderValues)readingOrder.Value : null
        };
    }

    partial void OnDxfParsed(XlsxDifferentialFontValue? font, (int NumFmtId, XlsxNumberFormat Format)? numFmt, XlsxFillFormatValue? fill, XlsxDifferentialAlignmentValue? alignment, XlsxDifferentialBorderValue? border, XlsxDifferentialProtectionValue? protection)
    {
        var dxf = new XlsxDxfValue
        {
            NumberFormat = numFmt?.Format,
            Font = font ?? XlsxDifferentialFontValue.Empty,
            Fill = fill is not null ? new XlsxDifferentialFillValue(fill) : XlsxDifferentialFillValue.Empty,
            Alignment = alignment ?? XlsxDifferentialAlignmentValue.Empty,
            Border = border ?? XlsxDifferentialBorderValue.Empty,
            Protection = protection ?? XlsxDifferentialProtectionValue.Empty,
        };
        _styles.AddDifferentialFormat(dxf);
    }

    partial void OnTableStyleElementParsed((TS?, PTS?) type, uint size, uint? dxfId)
    {
        // Skip definition without a differential format
        if (dxfId is null)
            return;

        // Excel permits only 0-9
        if (size > 9)
            throw PartStructureException.InvalidAttributeFormat(nameof(size), size.ToString(), _reader);

        // If there is a duplicate definition for a type, last one wins
        var dxf = _styles.DifferentialFormats[checked((int)dxfId.Value)];

        if (type.Item1 is { } tableStyleRegion)
            _currentTableStyle[tableStyleRegion] = (dxf, (int)size);

        if (type.Item2 is { } pivotStyleRegion)
            _currentPivotStyle[pivotStyleRegion] = (dxf, (int)size);
    }

    partial void OnTableStyleParsed(string name, bool pivot, bool table, uint? count)
    {
        // Because of tableStyle element duality, we are filling both styles and
        // only insert types that have set flag.
        if (table)
        {
            var tableStyle = new XlsxTableTheme(name);
            foreach (var (region, (dxf, bandSize)) in _currentTableStyle)
                tableStyle.SetRegionFormat(region, dxf, bandSize);

            _styles.AddTableStyle(tableStyle);
        }

        _currentTableStyle = new Dictionary<TS, (XlsxDxfValue Dxf, int BandSize)>();

        if (pivot)
        {
            var pivotStyle = new XlsxPivotTableStyle(name);
            foreach (var (region, (dxf, bandSize)) in _currentPivotStyle)
                pivotStyle.SetRegionFormat(region, dxf, bandSize);

            _styles.AddPivotStyle(pivotStyle);
        }

        _currentPivotStyle = new Dictionary<PTS, (XlsxDxfValue Dxf, int BandSize)>();
    }

    partial void OnTableStylesParsed(uint? count, string? defaultTableStyle, string? defaultPivotStyle)
    {
        if (!string.IsNullOrEmpty(defaultTableStyle))
            _styles.DefaultTableStyle = defaultTableStyle;

        if (!string.IsNullOrEmpty(defaultPivotStyle))
            _styles.DefaultPivotStyle = defaultPivotStyle;
    }

    private uint OnRgbColorParsed(uint? rgb)
    {
        // Despite the name, it's ARGB. If not specified, use black (Excel supplies 0x00000000, but
        // Excel plays very fast and loose with transparency).
        return rgb ?? 0xFF000000;
    }

    partial void OnIndexedColorsParsed(List<uint> rgbColor)
    {
        _styles.SetIndexedColors(rgbColor);
    }

    partial void OnMRUColorsParsed(List<XlsxColor> color)
    {
        _styles.SetMruColors(color);
    }

    private Xpr<XlsxColor> ParseColor(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxColor>();
        }

        return Xpr.From(_reader.ParseColor(elementName, ns));
    }

    private Xpr ParseExtensionList(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        _reader.Skip(elementName);
        return Xpr.Success();
    }

    private XlsxDifferentialProtectionValue OnCellProtectionParsed(bool? locked, bool? hidden)
    {
        return new XlsxDifferentialProtectionValue
        {
            Locked = locked,
            Hidden = hidden
        };
    }

    private XlsxFillFormatValue OnFillPatternFillParsed(XlsxFillFormatValue patternFillValue)
    {
        return patternFillValue;
    }

    private XlsxFillFormatValue OnFillGradientFillParsed(XlsxFillFormatValue gradientFillValue)
    {
        return gradientFillValue;
    }

    /// <summary>
    /// A mapping of <c>ST_TableStyleType</c>. Custom enum mapping due to table/pivot duality.
    /// </summary>
    private static readonly Dictionary<string, (TS?, PTS?)> TableStyleTypeMap = new()
    {
        { "wholeTable", (TS.WholeTable, PTS.WholeTable) },
        { "headerRow", (TS.HeaderRow, PTS.HeaderRow) },
        { "totalRow", (TS.TotalRow, PTS.GrandTotalRow) },
        { "firstColumn", (TS.FirstColumn, PTS.FirstColumn) },
        { "lastColumn", (TS.LastColumn, PTS.GrandTotalColumn) },
        { "firstRowStripe", (TS.FirstRowStripe, PTS.FirstRowStripe) },
        { "secondRowStripe", (TS.SecondRowStripe, PTS.SecondRowStripe) },
        { "firstColumnStripe", (TS.FirstColumnStripe, PTS.FirstColumnStripe) },
        { "secondColumnStripe", (TS.SecondColumnStripe, PTS.SecondColumnStripe) },
        { "firstHeaderCell", (TS.FirstHeaderCell, PTS.FirstHeaderCell) },
        { "lastHeaderCell", (TS.LastHeaderCell, null) },
        { "firstTotalCell", ( TS.FirstTotalCell, null) },
        { "lastTotalCell", ( TS.LastTotalCell, null) },
        { "firstSubtotalColumn", (null, PTS.SubtotalColumn1) },
        { "secondSubtotalColumn", (null,PTS.SubtotalColumn2) },
        { "thirdSubtotalColumn", (null, PTS.SubtotalColumn3) },
        { "firstSubtotalRow", (null, PTS.SubtotalRow1) },
        { "secondSubtotalRow", (null, PTS.SubtotalRow2) },
        { "thirdSubtotalRow", (null, PTS.SubtotalRow3) },
        { "blankRow", (null, PTS.BlankRow) },
        { "firstColumnSubheading", (null, PTS.ColumnSubheading1) },
        { "secondColumnSubheading", (null, PTS.ColumnSubheading2) },
        { "thirdColumnSubheading", (null, PTS.ColumnSubheading3) },
        { "firstRowSubheading", (null, PTS.RowSubheading1) },
        { "secondRowSubheading", (null, PTS.RowSubheading2) },
        { "thirdRowSubheading", (null, PTS.RowSubheading3) },
        { "pageFieldLabels", (null, PTS.PageFieldLabels) },
        { "pageFieldValues", (null, PTS.PageFieldValues) },
    };
}
