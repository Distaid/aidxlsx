#nullable enable

using System.Collections.Generic;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Opc;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Internal.IO;

internal partial class StylesReader
{
    private Xpr ParseNumFmts(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        OnNumFmtsParsing(count);

        var numFmt = new List<(int NumFmtId, XlsxNumberFormat Format)>();
        while (ParseNumFmt("numFmt", _ns) is { IsSuccess: true} numFmtItem)
        {
            numFmt.Add(numFmtItem.Value);
        }
        _reader.Close(elementName, ns);

        OnNumFmtsParsed(numFmt, count);
        return Xpr.Success();
    }

    partial void OnNumFmtsParsing(uint? count);

    partial void OnNumFmtsParsed(List<(int NumFmtId, XlsxNumberFormat Format)> numFmt, uint? count);

    private Xpr<(int NumFmtId, XlsxNumberFormat Format)> ParseNumFmt(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(int NumFmtId, XlsxNumberFormat Format)>();
        }

        var numFmtId = _reader.GetUInt("numFmtId");
        var formatCode = _reader.GetXString("formatCode");

        OnNumFmtParsing(numFmtId, formatCode);

        _reader.Close(elementName, ns);

        return Xpr.From(OnNumFmtParsed(numFmtId, formatCode));
    }

    partial void OnNumFmtParsing(uint numFmtId, string formatCode);

    private Xpr ParseFonts(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        OnFontsParsing(count);

        var font = new List<XlsxDifferentialFontValue>();
        while (ParseFont("font", _ns) is { IsSuccess: true} fontItem)
        {
            font.Add(fontItem.Value);
        }
        _reader.Close(elementName, ns);

        OnFontsParsed(font, count);
        return Xpr.Success();
    }

    partial void OnFontsParsing(uint? count);

    partial void OnFontsParsed(List<XlsxDifferentialFontValue> font, uint? count);

    private Xpr ParseFills(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        OnFillsParsing(count);

        var fill = new List<XlsxFillFormatValue>();
        while (ParseFill("fill", _ns) is { IsSuccess: true} fillItem)
        {
            fill.Add(fillItem.Value);
        }
        _reader.Close(elementName, ns);

        OnFillsParsed(fill, count);
        return Xpr.Success();
    }

    partial void OnFillsParsing(uint? count);

    partial void OnFillsParsed(List<XlsxFillFormatValue> fill, uint? count);

    private Xpr<XlsxFillFormatValue> ParseFill(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxFillFormatValue>();
        }

        OnFillParsing();

        XlsxFillFormatValue? choice;
        if (ParsePatternFill("patternFill", _ns) is { IsSuccess: true } patternFill)
        {
            choice = OnFillPatternFillParsed(patternFill.Value);
        }
        else if (ParseGradientFill("gradientFill", _ns) is { IsSuccess: true } gradientFill)
        {
            choice = OnFillGradientFillParsed(gradientFill.Value);
        }
        else
        {
            choice = default;
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnFillParsed(choice));
    }

    partial void OnFillParsing();

    private Xpr<XlsxFillFormatValue> ParsePatternFill(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxFillFormatValue>();
        }

        var patternType = _reader.GetOptionalEnum<XlsxFillPatternValues>("patternType");

        OnPatternFillParsing(patternType);

        var fgColorResult = ParseColor("fgColor", _ns);
        var fgColor = fgColorResult.IsSuccess ? fgColorResult.Value : default(XlsxColor?);
        var bgColorResult = ParseColor("bgColor", _ns);
        var bgColor = bgColorResult.IsSuccess ? bgColorResult.Value : default(XlsxColor?);
        _reader.Close(elementName, ns);

        return Xpr.From(OnPatternFillParsed(fgColor, bgColor, patternType));
    }

    partial void OnPatternFillParsing(XlsxFillPatternValues? patternType);

    private Xpr<XlsxFillFormatValue> ParseGradientFill(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxFillFormatValue>();
        }

        var type = _reader.GetOptionalEnum<XlsxGradientType>("type") ?? XlsxGradientType.Linear;
        var degree = _reader.GetOptionalDouble("degree") ?? 0;
        var left = _reader.GetOptionalDouble("left") ?? 0;
        var right = _reader.GetOptionalDouble("right") ?? 0;
        var top = _reader.GetOptionalDouble("top") ?? 0;
        var bottom = _reader.GetOptionalDouble("bottom") ?? 0;

        OnGradientFillParsing(type, degree, left, right, top, bottom);

        var stop = new List<(FractionOfOne Value, XlsxColor Color)>();
        while (ParseGradientStop("stop", _ns) is { IsSuccess: true} stopItem)
        {
            stop.Add(stopItem.Value);
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnGradientFillParsed(stop, type, degree, left, right, top, bottom));
    }

    partial void OnGradientFillParsing(XlsxGradientType type, double degree, double left, double right, double top, double bottom);

    private Xpr<(FractionOfOne Value, XlsxColor Color)> ParseGradientStop(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(FractionOfOne Value, XlsxColor Color)>();
        }

        var position = _reader.GetDouble("position");

        OnGradientStopParsing(position);

        var color = ParseColor("color", _ns).Value;
        _reader.Close(elementName, ns);

        return Xpr.From(OnGradientStopParsed(color, position));
    }

    partial void OnGradientStopParsing(double position);

    private Xpr ParseBorders(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        OnBordersParsing(count);

        var border = new List<XlsxDifferentialBorderValue>();
        while (ParseBorder("border", _ns) is { IsSuccess: true} borderItem)
        {
            border.Add(borderItem.Value);
        }
        _reader.Close(elementName, ns);

        OnBordersParsed(border, count);
        return Xpr.Success();
    }

    partial void OnBordersParsing(uint? count);

    partial void OnBordersParsed(List<XlsxDifferentialBorderValue> border, uint? count);

    private Xpr<XlsxDifferentialBorderValue> ParseBorder(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxDifferentialBorderValue>();
        }

        var diagonalUp = _reader.GetOptionalBool("diagonalUp");
        var diagonalDown = _reader.GetOptionalBool("diagonalDown");
        var outline = _reader.GetOptionalBool("outline") ?? true;

        OnBorderParsing(diagonalUp, diagonalDown, outline);

        var startResult = ParseBorderPr("start", _ns);
        var start = startResult.IsSuccess ? startResult.Value : default(XlsxBorderLine?);
        var endResult = ParseBorderPr("end", _ns);
        var end = endResult.IsSuccess ? endResult.Value : default(XlsxBorderLine?);
        var leftResult = ParseBorderPr("left", _ns);
        var left = leftResult.IsSuccess ? leftResult.Value : default(XlsxBorderLine?);
        var rightResult = ParseBorderPr("right", _ns);
        var right = rightResult.IsSuccess ? rightResult.Value : default(XlsxBorderLine?);
        var topResult = ParseBorderPr("top", _ns);
        var top = topResult.IsSuccess ? topResult.Value : default(XlsxBorderLine?);
        var bottomResult = ParseBorderPr("bottom", _ns);
        var bottom = bottomResult.IsSuccess ? bottomResult.Value : default(XlsxBorderLine?);
        var diagonalResult = ParseBorderPr("diagonal", _ns);
        var diagonal = diagonalResult.IsSuccess ? diagonalResult.Value : default(XlsxBorderLine?);
        var verticalResult = ParseBorderPr("vertical", _ns);
        var vertical = verticalResult.IsSuccess ? verticalResult.Value : default(XlsxBorderLine?);
        var horizontalResult = ParseBorderPr("horizontal", _ns);
        var horizontal = horizontalResult.IsSuccess ? horizontalResult.Value : default(XlsxBorderLine?);
        _reader.Close(elementName, ns);

        return Xpr.From(OnBorderParsed(start, end, left, right, top, bottom, diagonal, vertical, horizontal, diagonalUp, diagonalDown, outline));
    }

    partial void OnBorderParsing(bool? diagonalUp, bool? diagonalDown, bool outline);

    private Xpr<XlsxBorderLine> ParseBorderPr(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxBorderLine>();
        }

        var style = _reader.GetOptionalEnum<XlsxBorderStyleValues>("style") ?? XlsxBorderStyleValues.None;

        OnBorderPrParsing(style);

        var colorResult = ParseColor("color", _ns);
        var color = colorResult.IsSuccess ? colorResult.Value : default(XlsxColor?);
        _reader.Close(elementName, ns);

        return Xpr.From(OnBorderPrParsed(color, style));
    }

    partial void OnBorderPrParsing(XlsxBorderStyleValues style);

    private Xpr ParseCellStyleXfs(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        OnCellStyleXfsParsing(count);

        var xf = new List<(XlsxCellFormatValue Format, int? CellStyleXfId)>();
        while (ParseXf("xf", _ns) is { IsSuccess: true} xfItem)
        {
            xf.Add(xfItem.Value);
        }

        if (xf.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnCellStyleXfsParsed(xf, count);
        return Xpr.Success();
    }

    partial void OnCellStyleXfsParsing(uint? count);

    partial void OnCellStyleXfsParsed(List<(XlsxCellFormatValue Format, int? CellStyleXfId)> xf, uint? count);

    private Xpr<(XlsxCellFormatValue Format, int? CellStyleXfId)> ParseXf(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(XlsxCellFormatValue Format, int? CellStyleXfId)>();
        }

        var numFmtId = _reader.GetOptionalUInt("numFmtId");
        var fontId = _reader.GetOptionalUInt("fontId");
        var fillId = _reader.GetOptionalUInt("fillId");
        var borderId = _reader.GetOptionalUInt("borderId");
        var xfId = _reader.GetOptionalUInt("xfId");
        var quotePrefix = _reader.GetOptionalBool("quotePrefix") ?? false;
        var pivotButton = _reader.GetOptionalBool("pivotButton") ?? false;
        var applyNumberFormat = _reader.GetOptionalBool("applyNumberFormat");
        var applyFont = _reader.GetOptionalBool("applyFont");
        var applyFill = _reader.GetOptionalBool("applyFill");
        var applyBorder = _reader.GetOptionalBool("applyBorder");
        var applyAlignment = _reader.GetOptionalBool("applyAlignment");
        var applyProtection = _reader.GetOptionalBool("applyProtection");

        OnXfParsing(numFmtId, fontId, fillId, borderId, xfId, quotePrefix, pivotButton, applyNumberFormat, applyFont, applyFill, applyBorder, applyAlignment, applyProtection);

        var alignmentResult = ParseCellAlignment("alignment", _ns);
        var alignment = alignmentResult.IsSuccess ? alignmentResult.Value : default(XlsxDifferentialAlignmentValue?);
        var protectionResult = ParseCellProtection("protection", _ns);
        var protection = protectionResult.IsSuccess ? protectionResult.Value : default(XlsxDifferentialProtectionValue?);
        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnXfParsed(alignment, protection, numFmtId, fontId, fillId, borderId, xfId, quotePrefix, pivotButton, applyNumberFormat, applyFont, applyFill, applyBorder, applyAlignment, applyProtection));
    }

    partial void OnXfParsing(uint? numFmtId, uint? fontId, uint? fillId, uint? borderId, uint? xfId, bool quotePrefix, bool pivotButton, bool? applyNumberFormat, bool? applyFont, bool? applyFill, bool? applyBorder, bool? applyAlignment, bool? applyProtection);

    private Xpr<XlsxDifferentialAlignmentValue> ParseCellAlignment(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxDifferentialAlignmentValue>();
        }

        var horizontal = _reader.GetOptionalEnum<XlsxAlignmentHorizontalValues>("horizontal");
        var vertical = _reader.GetOptionalEnum<XlsxAlignmentVerticalValues>("vertical") ?? XlsxAlignmentVerticalValues.Bottom;
        var textRotation = _reader.GetOptionalUInt("textRotation");
        var wrapText = _reader.GetOptionalBool("wrapText");
        var indent = _reader.GetOptionalUInt("indent");
        var relativeIndent = _reader.GetOptionalInt("relativeIndent");
        var justifyLastLine = _reader.GetOptionalBool("justifyLastLine");
        var shrinkToFit = _reader.GetOptionalBool("shrinkToFit");
        var readingOrder = _reader.GetOptionalUInt("readingOrder");

        OnCellAlignmentParsing(horizontal, vertical, textRotation, wrapText, indent, relativeIndent, justifyLastLine, shrinkToFit, readingOrder);

        _reader.Close(elementName, ns);

        return Xpr.From(OnCellAlignmentParsed(horizontal, vertical, textRotation, wrapText, indent, relativeIndent, justifyLastLine, shrinkToFit, readingOrder));
    }

    partial void OnCellAlignmentParsing(XlsxAlignmentHorizontalValues? horizontal, XlsxAlignmentVerticalValues vertical, uint? textRotation, bool? wrapText, uint? indent, int? relativeIndent, bool? justifyLastLine, bool? shrinkToFit, uint? readingOrder);

    private Xpr<XlsxDifferentialProtectionValue> ParseCellProtection(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XlsxDifferentialProtectionValue>();
        }

        var locked = _reader.GetOptionalBool("locked");
        var hidden = _reader.GetOptionalBool("hidden");

        OnCellProtectionParsing(locked, hidden);

        _reader.Close(elementName, ns);

        return Xpr.From(OnCellProtectionParsed(locked, hidden));
    }

    partial void OnCellProtectionParsing(bool? locked, bool? hidden);

    private Xpr<List<(XlsxCellFormatValue Format, int? CellStyleXfId)>> ParseCellXfs(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<List<(XlsxCellFormatValue Format, int? CellStyleXfId)>>();
        }

        var count = _reader.GetOptionalUInt("count");

        OnCellXfsParsing(count);

        var xf = new List<(XlsxCellFormatValue Format, int? CellStyleXfId)>();
        while (ParseXf("xf", _ns) is { IsSuccess: true} xfItem)
        {
            xf.Add(xfItem.Value);
        }

        if (xf.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnCellXfsParsed(xf, count));
    }

    partial void OnCellXfsParsing(uint? count);

    private Xpr<Dictionary<int, XlsxCellStyleValue>> ParseCellStyles(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<Dictionary<int, XlsxCellStyleValue>>();
        }

        var count = _reader.GetOptionalUInt("count");

        OnCellStylesParsing(count);

        var cellStyle = new List<(int CellStyleXfId, XlsxCellStyleValue Style)>();
        while (ParseCellStyle("cellStyle", _ns) is { IsSuccess: true} cellStyleItem)
        {
            cellStyle.Add(cellStyleItem.Value);
        }

        if (cellStyle.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnCellStylesParsed(cellStyle, count));
    }

    partial void OnCellStylesParsing(uint? count);

    private Xpr<(int CellStyleXfId, XlsxCellStyleValue Style)> ParseCellStyle(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(int CellStyleXfId, XlsxCellStyleValue Style)>();
        }

        var name = _reader.GetOptionalXString("name");
        var xfId = _reader.GetUInt("xfId");
        var builtinId = _reader.GetOptionalUInt("builtinId");
        var iLevel = _reader.GetOptionalUInt("iLevel");
        var hidden = _reader.GetOptionalBool("hidden");
        var customBuiltin = _reader.GetOptionalBool("customBuiltin");

        OnCellStyleParsing(name, xfId, builtinId, iLevel, hidden, customBuiltin);

        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnCellStyleParsed(name, xfId, builtinId, iLevel, hidden, customBuiltin));
    }

    partial void OnCellStyleParsing(string? name, uint xfId, uint? builtinId, uint? iLevel, bool? hidden, bool? customBuiltin);

    private Xpr ParseDxfs(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");

        OnDxfsParsing(count);

        while (ParseDxf("dxf", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'dxf' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnDxfsParsed(count);
        return Xpr.Success();
    }

    partial void OnDxfsParsing(uint? count);

    partial void OnDxfsParsed(uint? count);

    private Xpr ParseDxf(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        OnDxfParsing();

        var fontResult = ParseFont("font", _ns);
        var font = fontResult.IsSuccess ? fontResult.Value : default(XlsxDifferentialFontValue?);
        var numFmtResult = ParseNumFmt("numFmt", _ns);
        var numFmt = numFmtResult.IsSuccess ? numFmtResult.Value : default((int NumFmtId, XlsxNumberFormat Format)?);
        var fillResult = ParseFill("fill", _ns);
        var fill = fillResult.IsSuccess ? fillResult.Value : default(XlsxFillFormatValue?);
        var alignmentResult = ParseCellAlignment("alignment", _ns);
        var alignment = alignmentResult.IsSuccess ? alignmentResult.Value : default(XlsxDifferentialAlignmentValue?);
        var borderResult = ParseBorder("border", _ns);
        var border = borderResult.IsSuccess ? borderResult.Value : default(XlsxDifferentialBorderValue?);
        var protectionResult = ParseCellProtection("protection", _ns);
        var protection = protectionResult.IsSuccess ? protectionResult.Value : default(XlsxDifferentialProtectionValue?);
        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        OnDxfParsed(font, numFmt, fill, alignment, border, protection);
        return Xpr.Success();
    }

    partial void OnDxfParsing();

    partial void OnDxfParsed(XlsxDifferentialFontValue? font, (int NumFmtId, XlsxNumberFormat Format)? numFmt, XlsxFillFormatValue? fill, XlsxDifferentialAlignmentValue? alignment, XlsxDifferentialBorderValue? border, XlsxDifferentialProtectionValue? protection);

    private Xpr ParseTableStyles(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var count = _reader.GetOptionalUInt("count");
        var defaultTableStyle = _reader.GetOptionalString("defaultTableStyle");
        var defaultPivotStyle = _reader.GetOptionalString("defaultPivotStyle");

        OnTableStylesParsing(count, defaultTableStyle, defaultPivotStyle);

        while (ParseTableStyle("tableStyle", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tableStyle' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnTableStylesParsed(count, defaultTableStyle, defaultPivotStyle);
        return Xpr.Success();
    }

    partial void OnTableStylesParsing(uint? count, string? defaultTableStyle, string? defaultPivotStyle);

    partial void OnTableStylesParsed(uint? count, string? defaultTableStyle, string? defaultPivotStyle);

    private Xpr ParseTableStyle(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var name = _reader.GetString("name");
        var pivot = _reader.GetOptionalBool("pivot") ?? true;
        var table = _reader.GetOptionalBool("table") ?? true;
        var count = _reader.GetOptionalUInt("count");

        OnTableStyleParsing(name, pivot, table, count);

        while (ParseTableStyleElement("tableStyleElement", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tableStyleElement' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnTableStyleParsed(name, pivot, table, count);
        return Xpr.Success();
    }

    partial void OnTableStyleParsing(string name, bool pivot, bool table, uint? count);

    partial void OnTableStyleParsed(string name, bool pivot, bool table, uint? count);

    private Xpr ParseTableStyleElement(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var type = _reader.GetStringMappedValue("type", TableStyleTypeMap);
        var size = _reader.GetOptionalUInt("size") ?? 1;
        var dxfId = _reader.GetOptionalUInt("dxfId");

        OnTableStyleElementParsing(type, size, dxfId);

        _reader.Close(elementName, ns);

        OnTableStyleElementParsed(type, size, dxfId);
        return Xpr.Success();
    }

    partial void OnTableStyleElementParsing((XlsxTableStyleRegionValues?, XlsxPivotStyleRegionValues?) type, uint size, uint? dxfId);

    partial void OnTableStyleElementParsed((XlsxTableStyleRegionValues?, XlsxPivotStyleRegionValues?) type, uint size, uint? dxfId);

    private Xpr ParseColors(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        OnColorsParsing();

        if (ParseIndexedColors("indexedColors", _ns) is { IsSuccess: true })
        {
            // Optional element 'indexedColors' was present
        }
        if (ParseMRUColors("mruColors", _ns) is { IsSuccess: true })
        {
            // Optional element 'mruColors' was present
        }
        _reader.Close(elementName, ns);

        OnColorsParsed();
        return Xpr.Success();
    }

    partial void OnColorsParsing();

    partial void OnColorsParsed();

    private Xpr ParseIndexedColors(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        OnIndexedColorsParsing();

        var rgbColor = new List<uint>();
        while (ParseRgbColor("rgbColor", _ns) is { IsSuccess: true} rgbColorItem)
        {
            rgbColor.Add(rgbColorItem.Value);
        }

        if (rgbColor.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnIndexedColorsParsed(rgbColor);
        return Xpr.Success();
    }

    partial void OnIndexedColorsParsing();

    partial void OnIndexedColorsParsed(List<uint> rgbColor);

    private Xpr ParseMRUColors(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        OnMRUColorsParsing();

        var color = new List<XlsxColor>();
        while (ParseColor("color", _ns) is { IsSuccess: true} colorItem)
        {
            color.Add(colorItem.Value);
        }

        if (color.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnMRUColorsParsed(color);
        return Xpr.Success();
    }

    partial void OnMRUColorsParsing();

    partial void OnMRUColorsParsed(List<XlsxColor> color);

    private Xpr<uint> ParseRgbColor(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<uint>();
        }

        var rgb = _reader.GetOptionalUIntHex("rgb");

        OnRgbColorParsing(rgb);

        _reader.Close(elementName, ns);

        return Xpr.From(OnRgbColorParsed(rgb));
    }

    partial void OnRgbColorParsing(uint? rgb);
}
