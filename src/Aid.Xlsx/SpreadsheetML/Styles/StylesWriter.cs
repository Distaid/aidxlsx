using System.Xml;
using PivotRegion = Aid.Xlsx.Model.Styles.Formatting.XlsxPivotStyleRegionValues;
using TableRegion = Aid.Xlsx.Model.Styles.Formatting.XlsxTableStyleRegionValues;
using static Aid.Xlsx.SpreadsheetML.Xml.OpenXmlConst;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Opc.Xml;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.SpreadsheetML.Styles;

internal class StylesWriter
{
    private const int FirstUserDefinedFormatIndex = XlsxWorkbookStyles.FirstUserDefinedNumberFormatIndex;

    private static readonly List<(string Type, TableRegion? TableRegion, PivotRegion? PivotRegion)> TableRegionsMap = new()
    {
        ("wholeTable", TableRegion.WholeTable, PivotRegion.WholeTable),
        ("headerRow", TableRegion.HeaderRow, PivotRegion.HeaderRow),
        ("totalRow", TableRegion.TotalRow, PivotRegion.GrandTotalRow),
        ("firstColumn", TableRegion.FirstColumn, PivotRegion.FirstColumn),
        ("lastColumn", TableRegion.LastColumn, PivotRegion.GrandTotalColumn),
        ("firstRowStripe", TableRegion.FirstRowStripe, PivotRegion.FirstRowStripe),
        ("secondRowStripe", TableRegion.SecondRowStripe, PivotRegion.SecondRowStripe),
        ("firstColumnStripe", TableRegion.FirstColumnStripe, PivotRegion.FirstColumnStripe),
        ("secondColumnStripe", TableRegion.SecondColumnStripe, PivotRegion.SecondColumnStripe),
        ("firstHeaderCell", TableRegion.FirstHeaderCell, PivotRegion.FirstHeaderCell),
        ("lastHeaderCell", TableRegion.LastHeaderCell, null),
        ("firstTotalCell", TableRegion.FirstTotalCell, null),
        ("lastTotalCell", TableRegion.LastTotalCell, null),
        ("firstSubtotalColumn", null, PivotRegion.SubtotalColumn1),
        ("secondSubtotalColumn", null, PivotRegion.SubtotalColumn2),
        ("thirdSubtotalColumn", null, PivotRegion.SubtotalColumn3),
        ("firstSubtotalRow", null, PivotRegion.SubtotalRow1),
        ("secondSubtotalRow", null, PivotRegion.SubtotalRow2),
        ("thirdSubtotalRow", null, PivotRegion.SubtotalRow3),
        ("blankRow", null, PivotRegion.BlankRow),
        ("firstColumnSubheading", null, PivotRegion.ColumnSubheading1),
        ("secondColumnSubheading", null, PivotRegion.ColumnSubheading2),
        ("thirdColumnSubheading", null, PivotRegion.ColumnSubheading3),
        ("firstRowSubheading", null, PivotRegion.RowSubheading1),
        ("secondRowSubheading", null, PivotRegion.RowSubheading2),
        ("thirdRowSubheading", null, PivotRegion.RowSubheading3),
        ("pageFieldLabels", null, PivotRegion.PageFieldLabels),
        ("pageFieldValues", null, PivotRegion.PageFieldValues),
    };

    private readonly string _ns = Main2006SsNs;

    internal void WriteContent(Stream partStream, IEnumMapper mapper, XlsxWorkbookStyles styles, XlsxWorkbook workbook, XlsxWorkbook.SaveContext context)
    {
        // Determine which format components are used and thus should be saved.
        var usedCellFormats = new HashSet<XlsxCellFormatValue>(ReferenceEqualityComparer<XlsxCellFormatValue>.Instance)
        {
            styles.DefaultFormat
        };
        foreach (var sheet in workbook.WorksheetsInternal)
        {
            if (sheet.FormatValue is { } sheetFormat)
                usedCellFormats.Add(sheetFormat);

            foreach (var column in sheet.Internals.ColumnsCollection.Values)
            {
                if (column.FormatValue is { } columnFormat)
                    usedCellFormats.Add(columnFormat);
            }

            foreach (var row in sheet.Internals.RowsCollection.Values)
            {
                if (row.FormatValue is { } rowFormat)
                    usedCellFormats.Add(rowFormat);
            }

            sheet.Internals.CellsCollection.FormatSlice.AddUsedFormat(usedCellFormats);
        }

        var usedNumberFormats = new HashSet<XlsxNumberFormat>();
        var usedFonts = new HashSet<XlsxFontFormatValue>();
        var usedFills = new HashSet<XlsxFillFormatValue>();
        var usedBorders = new HashSet<XlsxBorderFormatValue>();
        foreach (var cellFormat in usedCellFormats)
        {
            if (cellFormat.NumberFormat is { } numberFormat)
                usedNumberFormats.Add(numberFormat);

            if (cellFormat.Font is { } font)
                usedFonts.Add(font);

            if (cellFormat.Fill is { } fill)
                usedFills.Add(fill);

            if (cellFormat.Border is { } border)
                usedBorders.Add(border);
        }

        foreach (var cellStyle in styles.CellStyles.Values)
        {
            if (cellStyle.NumberFormat is { } numberFormat)
                usedNumberFormats.Add(numberFormat);

            if (cellStyle.Font is { } font)
                usedFonts.Add(font);

            if (cellStyle.Fill is { } fill)
                usedFills.Add(fill);

            if (cellStyle.Border is { } border)
                usedBorders.Add(border);
        }

        // SST writes fonts as ids, unlike other format properties which are inlined next to a rich text
        foreach (var phoneticsFont in workbook.SharedStringTable.GetUsedPhoneticFonts())
            usedFonts.Add(phoneticsFont);

        // Create dxfMap from used dxfs in cfs, tables, pivot tables and so on
        var usedDxf = new HashSet<XlsxDxfValue>();
        foreach (var ws in workbook.WorksheetsInternal)
        {
            foreach (var cf in ws.ConditionalFormats)
            {
                if (cf.FormatValue is { } dxf)
                    usedDxf.Add(dxf);
            }

            // Table styles should be retained in a workbook, even if not used in a table.
            foreach (var tableStyle in styles.TableStyles.Values)
            {
                foreach (var regionDxf in tableStyle.RegionFormats.Values)
                {
                    usedDxf.Add(regionDxf);
                }
            }

            foreach (var table in ws.Tables)
            {
                foreach (var field in table.Fields)
                {
                    if (field.HeaderFormatValue is { } headerDxf)
                        usedDxf.Add(headerDxf);

                    if (field.DataFormatValue is { } dataDxf)
                        usedDxf.Add(dataDxf);

                    if (field.TotalFormatValue is { } totalsDxf)
                        usedDxf.Add(totalsDxf);
                }
            }

            foreach (var pt in ws.PivotTables)
            {
                foreach (var format in pt.Formats)
                {
                    if (format.FormatValue is { } dxFormat)
                        usedDxf.Add(dxFormat);
                }

                foreach (var cf in pt.ConditionalFormats)
                {
                    if (cf.Format.FormatValue is { } cfDxf)
                        usedDxf.Add(cfDxf);
                }

                foreach (var field in pt.PivotFields)
                {
                    if (field.NumberFormatValue is { } fieldNumFmt)
                        usedNumberFormats.Add(fieldNumFmt);
                }
            }
        }

        var settings = new XmlWriterSettings
        {
            Encoding = XlsxHelper.NoBomUTF8
        };

        using var xml = new XmlTreeWriter(XmlWriter.Create(partStream, settings), mapper);

        xml.WriteStartDocument("styleSheet", _ns);

        // Number formats
        // The map has predefined formats from index 0 and the user defined ones from 164 onward.
        // There is a gap between predefined formats.
        var predefinedNumberFormats = XlsxPredefinedFormat.NumberFormatIds;
        var numberFormatMap = SequentialMap<int, XlsxNumberFormat>.Create(usedNumberFormats, styles.NumberFormats, XlsxPredefinedFormat.NumberFormatIds, FirstUserDefinedFormatIndex);

        if (numberFormatMap.Count > predefinedNumberFormats.Count)
            WriteNumberFormats(xml, numberFormatMap, predefinedNumberFormats);

        // Fonts. Register default format font as font zero. The font zero is used for font name and size.
        var firstFontValues = new Dictionary<XlsxFontFormatValue, int> { { styles.DefaultFormat.Font, 0 } };
        var fontFormatsMap = SequentialMap<int, XlsxFontFormatValue>.Create(usedFonts, styles.Fonts, firstFontValues);
        if (fontFormatsMap.Count > 0)
            WriteFonts(xml, fontFormatsMap);

        // Fill 0 must be None and fill 1 must be Gray125, that is just an immutable fact of the universe.
        // Excel will ignore fills at 0/1 and will use None/Gray125. Write both fills whether they are used
        // or not.
        AddFillAsUsed(XlsxFillFormatValue.None);
        AddFillAsUsed(XlsxFillFormatValue.Gray125);
        var firstFillValues = new Dictionary<XlsxFillFormatValue, int>
        {
            { XlsxFillFormatValue.None, 0 },
            { XlsxFillFormatValue.Gray125, 1 },
        };
        var fillsFormatsMap = SequentialMap<int, XlsxFillFormatValue>.Create(usedFills, styles.Fills, firstFillValues);
        WriteFills(xml, fillsFormatsMap);

        var borderFormatsMap = SequentialMap<int, XlsxBorderFormatValue>.Create(usedBorders, styles.Borders);
        if (borderFormatsMap.Count > 0)
            WriteBorders(xml, borderFormatsMap);

        var cellStylesMap = new SequentialMap<StyleId, XlsxCellStyleValue>(styles.CellStyles);

        // All cell styles, regardless if they are used or not should be written to the file
        foreach (var styleId in styles.CellStyles.Keys)
            cellStylesMap.Add(styleId);

        cellStylesMap.Sort();
        if (cellStylesMap.Count > 0)
            WriteCellStyleXfs(xml, cellStylesMap, numberFormatMap, fontFormatsMap, fillsFormatsMap, borderFormatsMap);

        var firstCellXfsValues = new Dictionary<XlsxCellFormatValue, int> { { styles.DefaultFormat, 0 } };
        var cellXfsMap = SequentialMap<int, XlsxCellFormatValue>.Create(usedCellFormats, styles.CellFormats, firstCellXfsValues);
        WriteCellXfs(xml, cellXfsMap, numberFormatMap, fontFormatsMap, fillsFormatsMap, borderFormatsMap, cellStylesMap);

        if (cellStylesMap.Count > 0)
            WriteCellStyles(xml, cellStylesMap);

        var dxfMap = SequentialMap<int, XlsxDxfValue>.Create(usedDxf, styles.DifferentialFormats);
        WriteDxfs(xml, dxfMap, numberFormatMap.Count);

        var hasTableStyles = styles.TableStyles.Count > 0 ||
                             styles.PivotStyles.Count > 0 ||
                             styles.DefaultTableStyle is not null ||
                             styles.DefaultPivotStyle is not null;
        if (hasTableStyles)
            WriteTableStyles(xml, dxfMap, styles);

        WriteColors(xml, styles);

        WriteExtensions(xml);

        xml.WriteEndDocument();

        // Fill the maps used in other parts to determine saved id for a format
        foreach (var (numberFormatId, numberFormat) in numberFormatMap.GetActual())
            context.NumberFormatMap.TryAdd(numberFormat, numberFormatId);

        foreach (var (fontId, fontFormat) in fontFormatsMap.GetActual())
            context.FontMap.TryAdd(fontFormat, fontId);

        foreach (var (xfId, format) in cellXfsMap.GetActual())
            context.FormatMap.TryAdd(format, (uint)xfId);

        foreach (var (dxfId, dxf) in dxfMap.GetActual())
            context.DxfMap.TryAdd(dxf, (uint)dxfId);

        return;

        void AddFillAsUsed(XlsxFillFormatValue format)
        {
            if (!styles.Fills.ContainsValue(format))
                styles.AddFillFormat(format);

            usedFills.Add(format);
        }
    }

    private void WriteNumberFormats(XmlTreeWriter xml, SequentialMap<int, XlsxNumberFormat> idMap, IReadOnlyDictionary<XlsxNumberFormat, int> predefinedNumberFormats)
    {
        xml.WriteStartElement("numFmts", _ns);
        xml.WriteAttribute("count", idMap.Count - predefinedNumberFormats.Count);

        foreach (var (savedId, format) in idMap.GetActual())
        {
            // Do not write the predefined formats, Excel doesn't write that either
            if (predefinedNumberFormats.ContainsKey(format))
                continue;

            WriteNumFmt(xml, "numFmt", savedId, format);
        }

        xml.WriteEndElement(); // numFmts
    }

    private void WriteNumFmt(XmlTreeWriter xml, string elementName, int numFmtId, string format)
    {
        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttribute("numFmtId", numFmtId);
        xml.WriteAttribute("formatCode", format);
        xml.WriteEndElement();
    }

    private void WriteFonts(XmlTreeWriter xml, SequentialMap<int, XlsxFontFormatValue> idMap)
    {
        xml.WriteStartElement("fonts", _ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach (var (_, font) in idMap.GetActual())
            WriteFont(xml, "font", font);

        xml.WriteEndElement();
    }

    private void WriteFont(XmlTreeWriter xml, string elementName, XlsxFontFormatValue font)
    {
        // MS-OI29500 dictates font elements order.
        xml.WriteStartElement(elementName, _ns);

        WriteFlag("b", font.Bold);
        WriteFlag("i", font.Italic);
        WriteFlag("strike", font.Strikethrough);
        WriteFlag("condense", font.Condense);
        WriteFlag("extend", font.Extend);
        WriteFlag("outline", font.Outline);
        WriteFlag("shadow", font.Shadow);

        if (font.Underline != XlsxFontUnderlineValues.None)
            WriteFontUnderline(xml, font.Underline);

        if (font.VerticalAlignment != XlsxFontVerticalTextAlignmentValues.Baseline)
            WriteFontVerticalAlignment(xml, font.VerticalAlignment);

        WriteFontSize(xml, font.Size);

        if (!font.Color.IsAuto)
            xml.WriteColor("color", _ns, font.Color);

        WriteFontName(xml, font.Name);

        if (font.Family != XlsxFontFamilyNumberingValues.NotApplicable)
            WriteFontFamily(xml, font.Family);

        if (font.Charset != XlsxFontCharSet.Ansi)
            WriteFontCharset(xml, font.Charset);

        if (font.Scheme != XlsxFontScheme.None)
            WriteFontScheme(xml, font.Scheme);

        xml.WriteEndElement();
        return;

        void WriteFlag(string flagName, bool flag)
        {
            if (flag)
                xml.WriteBooleanProperty(flagName, true, _ns);
        }
    }

    private void WriteFont(XmlTreeWriter xml, string elementName, XlsxDifferentialFontValue font)
    {
        // MS-OI29500 dictates font elements order.
        xml.WriteStartElement(elementName, _ns);

        WriteFlag("b", font.Bold);
        WriteFlag("i", font.Italic);
        WriteFlag("strike", font.Strikethrough);
        WriteFlag("condense", font.Condense);
        WriteFlag("extend", font.Extend);
        WriteFlag("outline", font.Outline);
        WriteFlag("shadow", font.Shadow);

        if (font.Underline is { } underline && underline != XlsxFontUnderlineValues.None)
            WriteFontUnderline(xml, underline);

        if (font.VerticalAlignment is { } verticalAlignment && verticalAlignment != XlsxFontVerticalTextAlignmentValues.Baseline)
            WriteFontVerticalAlignment(xml, verticalAlignment);

        if (font.Size is { } size)
            WriteFontSize(xml, size);

        if (font.Color is { } color && !color.IsAuto)
            xml.WriteColor("color", _ns, color);

        if (font.Name is { } name)
            WriteFontName(xml, name);

        if (font.Family is { } family && family != XlsxFontFamilyNumberingValues.NotApplicable)
            WriteFontFamily(xml, family);

        if (font.Charset is { } charset && charset != XlsxFontCharSet.Ansi)
            WriteFontCharset(xml, charset);

        if (font.Scheme is { } scheme && scheme != XlsxFontScheme.None)
            WriteFontScheme(xml, scheme);

        xml.WriteEndElement();
        return;

        void WriteFlag(string flagName, bool? flag)
        {
            if (flag == true)
                xml.WriteBooleanProperty(flagName, true, _ns);
        }
    }

    private void WriteFontUnderline(XmlTreeWriter xml, XlsxFontUnderlineValues underline)
    {
        xml.WriteStartElement("u", _ns);
        xml.WriteAttributeDefault("val", underline, XlsxFontUnderlineValues.Single);
        xml.WriteEndElement();
    }

    private void WriteFontVerticalAlignment(XmlTreeWriter xml, XlsxFontVerticalTextAlignmentValues verticalAlignment)
    {
        xml.WriteStartElement("vertAlign", _ns);
        xml.WriteAttribute("val", verticalAlignment);
        xml.WriteEndElement();
    }

    private void WriteFontSize(XmlTreeWriter xml, XlsxFontSize size)
    {
        xml.WriteStartElement("sz", _ns);
        xml.WriteAttribute("val", size.Points);
        xml.WriteEndElement();
    }

    private void WriteFontName(XmlTreeWriter xml, XlsxFontName fontName)
    {
        xml.WriteStartElement("name", _ns);
        xml.WriteAttribute("val", fontName.Text);
        xml.WriteEndElement();
    }

    private void WriteFontFamily(XmlTreeWriter xml, XlsxFontFamilyNumberingValues family)
    {
        xml.WriteStartElement("family", _ns);
        xml.WriteAttribute("val", (int)family);
        xml.WriteEndElement();
    }

    private void WriteFontCharset(XmlTreeWriter xml, XlsxFontCharSet charset)
    {
        // Charset is stored as an CT_IntProperty
        xml.WriteStartElement("charset", _ns);
        xml.WriteAttribute("val", (int)charset);
        xml.WriteEndElement();
    }

    private void WriteFontScheme(XmlTreeWriter xml, XlsxFontScheme scheme)
    {
        xml.WriteStartElement("scheme", _ns);
        xml.WriteAttribute("val", scheme);
        xml.WriteEndElement();
    }

    private void WriteFills(XmlTreeWriter xml, SequentialMap<int, XlsxFillFormatValue> idMap)
    {
        xml.WriteStartElement("fills", _ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach (var (_, fill) in idMap.GetActual())
            WriteFill(xml, "fill", fill.Pattern, fill.LinearGradient, fill.PathGradient, false);

        xml.WriteEndElement();
    }

    private void WriteFill(XmlTreeWriter xml, string elementName, XlsxPatternFill? patternFill, XlsxLinearGradientFill? linearGradient, XlsxPathGradientFill? pathGradient, bool isDxf)
    {
        xml.WriteStartElement(elementName, _ns);

        // A fill element with no pattern/gradient is a valid state per XML
        if (patternFill is not null)
        {
            xml.WriteStartElement("patternFill", _ns);
            xml.WriteAttribute("patternType", patternFill.PatternType);

            var patternColor = patternFill.PatternColor;
            var bgColor = patternFill.BackgroundColor;

            // Fix solid pattern discrepancy. The GUI shows solid fill color in the background
            // color picker, so it would be expected that it is stored in bgColor.
            // * internal structures store solid fill color in the 'IXlsxFill.BackgroundColor'
            // * For cell format, the 'patternFill' element stores it in the *fgColor*, not bgColor
            // * For dxf, the 'patternFill' stores it in the *bgColor*
            if (!isDxf && patternFill.PatternType == XlsxFillPatternValues.Solid)
            {
                (patternColor, bgColor) = (bgColor, patternColor);
            }

            if (patternColor.HasValue)
                xml.WriteColor("fgColor", _ns, patternColor);

            if (bgColor.HasValue)
                xml.WriteColor("bgColor", _ns, bgColor);

            xml.WriteEndElement();
        }
        else if (linearGradient is not null)
        {
            // Linear is the default type, so no need to write it
            xml.WriteStartElement("gradientFill", _ns);
            xml.WriteAttributeDefault("degree", linearGradient.Degrees, 0);

            WriteStops(linearGradient.Stops);

            xml.WriteEndElement();
        }
        else if (pathGradient is not null)
        {
            xml.WriteStartElement("gradientFill", _ns);
            xml.WriteAttribute("type", "path");
            xml.WriteAttributeDefault("left", pathGradient.InnerLeft.Value, 0);
            xml.WriteAttributeDefault("right", pathGradient.InnerRight.Value, 0);
            xml.WriteAttributeDefault("top", pathGradient.InnerTop.Value, 0);
            xml.WriteAttributeDefault("bottom", pathGradient.InnerBottom.Value, 0);

            WriteStops(pathGradient.Stops);

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        void WriteStops(IReadOnlyDictionary<FractionOfOne, XlsxColor> stops)
        {
            // Excel doesn't care about stop order by positions, but sort anyway
            foreach (var (position, color) in stops.OrderBy(x => x.Key.Value))
            {
                xml.WriteStartElement("stop", _ns);
                xml.WriteAttribute("position", position.Value);
                xml.WriteColor("color", _ns, color);
                xml.WriteEndElement();
            }
        }
    }

    private void WriteBorders(XmlTreeWriter xml, SequentialMap<int, XlsxBorderFormatValue> idMap)
    {
        xml.WriteStartElement("borders", _ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach (var (_, border) in idMap.GetActual())
            WriteBorder(xml, "border", border);

        xml.WriteEndElement();
    }

    private void WriteBorder(XmlTreeWriter xml, string elementName, XlsxBorderFormatValue border)
    {
        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttributeDefault("diagonalUp", border.DiagonalUp, false);
        xml.WriteAttributeDefault("diagonalDown", border.DiagonalDown, false);

        // ISO should be "start"+"end", but Excel uses "left"+"right"
        WriteBorderPr(xml, "left", border.Left);
        WriteBorderPr(xml, "right", border.Right);
        WriteBorderPr(xml, "top", border.Top);
        WriteBorderPr(xml, "bottom", border.Bottom);
        WriteBorderPr(xml, "diagonal", border.Diagonal);

        xml.WriteEndElement();
    }

    private void WriteBorder(XmlTreeWriter xml, string elementName, XlsxDifferentialBorderValue border)
    {
        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttributeDefault("diagonalUp", border.DiagonalUp, false);
        xml.WriteAttributeDefault("diagonalDown", border.DiagonalDown, false);

        // Outline has no meaning for cell styles, it is only for dxf - tables and such.
        xml.WriteAttributeDefault("outline", border.Outline, true);

        // ISO should be "start"+"end", but Excel uses "left"+"right"
        if (border.Left is { } left)
            WriteBorderPr(xml, "left", left);

        if (border.Right is { } right)
            WriteBorderPr(xml, "right", right);

        if (border.Top is { } top)
            WriteBorderPr(xml, "top", top);

        if (border.Bottom is { } bottom)
            WriteBorderPr(xml, "bottom", bottom);

        if (border.Diagonal is { } diagonal)
            WriteBorderPr(xml, "diagonal", diagonal);

        if (border.Vertical is { } vertical)
            WriteBorderPr(xml, "vertical", vertical);

        if (border.Horizontal is { } horizontal)
            WriteBorderPr(xml, "horizontal", horizontal);

        xml.WriteEndElement();
    }

    private void WriteBorderPr(XmlTreeWriter xml, string name, XlsxBorderLine borderLine)
    {
        xml.WriteStartElement(name, _ns);
        xml.WriteAttributeDefault("style", borderLine.Style, XlsxBorderStyleValues.None);
        if (borderLine.Style != XlsxBorderStyleValues.None)
        {
            // Color for border is always written and default value is automatic color.
            xml.WriteColor("color", _ns, borderLine.Color);
        }

        xml.WriteEndElement();
    }

    private void WriteCellStyleXfs(
        XmlTreeWriter xml,
        SequentialMap<StyleId, XlsxCellStyleValue> cellStylesMap,
        SequentialMap<int, XlsxNumberFormat> numFmtIdMap,
        SequentialMap<int, XlsxFontFormatValue> fontIdMap,
        SequentialMap<int, XlsxFillFormatValue> fillIdMap,
        SequentialMap<int, XlsxBorderFormatValue> borderIdMap)
    {
        xml.WriteStartElement("cellStyleXfs", _ns);
        xml.WriteAttribute("count", cellStylesMap.Count);

        // Collection must have at least one element
        foreach (var (_, cellStyle) in cellStylesMap.GetActual())
        {
            xml.WriteStartElement("xf", _ns);
            xml.WriteAttributeOptional("numFmtId", numFmtIdMap.GetSavedId(cellStyle.NumberFormat));
            xml.WriteAttributeOptional("fontId", fontIdMap.GetSavedId(cellStyle.Font));
            xml.WriteAttributeOptional("fillId", fillIdMap.GetSavedId(cellStyle.Fill));
            xml.WriteAttributeOptional("borderId", borderIdMap.GetSavedId(cellStyle.Border));

            // cellStyleXf doesn't use quote, pivot button or xfId -> skip those attributes
            xml.WriteAttributeDefault("applyNumberFormat", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.NumberFormat), true);
            xml.WriteAttributeDefault("applyFont", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Font), true);
            xml.WriteAttributeDefault("applyFill", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Fill), true);
            xml.WriteAttributeDefault("applyBorder", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Border), true);
            xml.WriteAttributeDefault("applyAlignment", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Alignment), true);
            xml.WriteAttributeDefault("applyProtection", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Protection), true);

            if (cellStyle.Alignment is { } alignment)
                WriteAlignment(xml, "alignment", alignment);

            if (cellStyle.Protection is { } protection)
                WriteProtection(xml, "protection", protection);

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteCellXfs(
        XmlTreeWriter xml,
        SequentialMap<int, XlsxCellFormatValue> idMap,
        SequentialMap<int, XlsxNumberFormat> numFmtIdMap,
        SequentialMap<int, XlsxFontFormatValue> fontIdMap,
        SequentialMap<int, XlsxFillFormatValue> fillIdMap,
        SequentialMap<int, XlsxBorderFormatValue> borderIdMap,
        SequentialMap<StyleId, XlsxCellStyleValue> cellStyleIdMap)
    {
        xml.WriteStartElement("cellXfs", _ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach (var (_, cellXf) in idMap.GetActual())
        {
            xml.WriteStartElement("xf", _ns);

            xml.WriteAttributeOptional("numFmtId", numFmtIdMap.GetSavedId(cellXf.NumberFormat));
            xml.WriteAttributeOptional("fontId", fontIdMap.GetSavedId(cellXf.Font));
            xml.WriteAttributeOptional("fillId", fillIdMap.GetSavedId(cellXf.Fill));
            xml.WriteAttributeOptional("borderId", borderIdMap.GetSavedId(cellXf.Border));

            if (cellXf.CellStyleId is not null)
                xml.WriteAttribute("xfId", cellStyleIdMap.GetSavedId(cellXf.CellStyleId.Value));

            xml.WriteAttributeDefault("quotePrefix", cellXf.IncludeQuotePrefix, false);
            xml.WriteAttributeDefault("pivotButton", cellXf.PivotButton, false);
            xml.WriteAttributeDefault("applyNumberFormat", cellXf.CustomFormat.HasFlag(CellFormatComponents.NumberFormat), false);
            xml.WriteAttributeDefault("applyFont", cellXf.CustomFormat.HasFlag(CellFormatComponents.Font), false);
            xml.WriteAttributeDefault("applyFill", cellXf.CustomFormat.HasFlag(CellFormatComponents.Fill), false);
            xml.WriteAttributeDefault("applyBorder", cellXf.CustomFormat.HasFlag(CellFormatComponents.Border), false);
            xml.WriteAttributeDefault("applyAlignment", cellXf.CustomFormat.HasFlag(CellFormatComponents.Alignment), false);
            xml.WriteAttributeDefault("applyProtection", cellXf.CustomFormat.HasFlag(CellFormatComponents.Protection), false);

            if (cellXf.Alignment is { } alignment)
                WriteAlignment(xml, "alignment", alignment);

            if (cellXf.Protection is { } protection)
                WriteProtection(xml, "protection", protection);

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteAlignment(XmlTreeWriter xml, string elementName, XlsxAlignmentFormatValue alignment)
    {
        if (alignment == XlsxAlignmentFormatValue.Default)
            return;

        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttributeDefault("horizontal", alignment.Horizontal, XlsxAlignmentFormatValue.Default.Horizontal);
        xml.WriteAttributeDefault("vertical", alignment.Vertical, XlsxAlignmentFormatValue.Default.Vertical);
        xml.WriteAttributeDefault("textRotation", alignment.TextRotation.GetIso(), XlsxAlignmentFormatValue.Default.TextRotation.Value);
        xml.WriteAttributeDefault("wrapText", alignment.WrapText, XlsxAlignmentFormatValue.Default.WrapText);
        xml.WriteAttributeDefault("indent", alignment.Indent, XlsxAlignmentFormatValue.Default.Indent);
        xml.WriteAttributeDefault("relativeIndent", alignment.RelativeIndent, XlsxAlignmentFormatValue.Default.RelativeIndent);
        xml.WriteAttributeDefault("justifyLastLine", alignment.JustifyLastLine, XlsxAlignmentFormatValue.Default.JustifyLastLine);
        xml.WriteAttributeDefault("shrinkToFit", alignment.ShrinkToFit, XlsxAlignmentFormatValue.Default.ShrinkToFit);
        xml.WriteAttributeDefault("readingOrder", (uint)alignment.ReadingOrder, (uint)XlsxAlignmentFormatValue.Default.ReadingOrder);
        xml.WriteEndElement();
    }

    private void WriteAlignment(XmlTreeWriter xml, string elementName, XlsxDifferentialAlignmentValue alignment)
    {
        // Alignment is kind of wonky. Current Excel doesn't even support it in a DXF dialogue and doesn't always work correctly.
        xml.WriteStartElement(elementName, _ns);
        if (alignment.Horizontal is { } horizontal)
            xml.WriteAttribute("horizontal", horizontal);

        if (alignment.Vertical is { } vertical)
            xml.WriteAttribute("vertical", vertical);

        if (alignment.TextRotation is { } textRotation)
            xml.WriteAttribute("textRotation", textRotation.GetIso());

        if (alignment.WrapText is { } wrapText)
            xml.WriteAttribute("wrapText", wrapText);

        if (alignment.Indent is { } indent)
            xml.WriteAttribute("indent", indent);

        if (alignment.RelativeIndent is { } relativeIndent)
            xml.WriteAttribute("relativeIndent", relativeIndent);

        if (alignment.JustifyLastLine is { } justifyLastLine)
            xml.WriteAttribute("justifyLastLine", justifyLastLine);

        if (alignment.ShrinkToFit is { } shrinkToFit)
            xml.WriteAttribute("shrinkToFit", shrinkToFit);

        if (alignment.ReadingOrder is { } readingOrder)
            xml.WriteAttribute("readingOrder", readingOrder);

        xml.WriteEndElement();
    }

    private void WriteProtection(XmlTreeWriter xml, string elementName, XlsxProtectionFormatValue protection)
    {
        if (protection.Locked && !protection.Hidden)
            return;

        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttributeDefault("locked", protection.Locked, true);
        xml.WriteAttributeDefault("hidden", protection.Hidden, false);
        xml.WriteEndElement();
    }

    private void WriteProtection(XmlTreeWriter xml, string elementName, XlsxDifferentialProtectionValue protection)
    {
        xml.WriteStartElement(elementName, _ns);

        if (protection.Locked is { } locked)
            xml.WriteAttribute("locked", locked);

        if (protection.Hidden is { } hidden)
            xml.WriteAttribute("hidden", hidden);

        xml.WriteEndElement();
    }

    private void WriteCellStyles(XmlTreeWriter xml, SequentialMap<StyleId, XlsxCellStyleValue> cellStylesMap)
    {
        xml.WriteStartElement("cellStyles", _ns);
        xml.WriteAttribute("count", cellStylesMap.Count);

        // Collection must have at least one element
        foreach (var (mappedStyleId, cellStyle) in cellStylesMap.GetActual())
        {
            xml.WriteStartElement("cellStyle", _ns);

            // Name is technically optional and Excel will generate one if missing, but we ensure the name always exist
            xml.WriteAttribute("name", cellStyle.Name);
            xml.WriteAttribute("xfId", mappedStyleId);
            if (cellStyle.BuiltInStyle is { } builtInStyle)
            {
                if (builtInStyle is >= BuiltInStyleValues.RowLevel1 and <= BuiltInStyleValues.RowLevel7)
                {
                    xml.WriteAttributeOptional("builtinId", 1);
                    xml.WriteAttribute("iLevel", BuiltInStyleValues.RowLevel1 - builtInStyle + 1);
                }
                else if (builtInStyle is >= BuiltInStyleValues.ColumnLevel1 and <= BuiltInStyleValues.ColumnLevel7)
                {
                    xml.WriteAttributeOptional("builtinId", 2);
                    xml.WriteAttribute("iLevel", BuiltInStyleValues.ColumnLevel1 - builtInStyle + 1);
                }
                else
                {
                    xml.WriteAttributeOptional("builtinId", (int?)cellStyle.BuiltInStyle);
                }
            }

            // Hidden + flag are optional per schema, but basically it's a bool with default
            xml.WriteAttributeDefault("hidden", cellStyle.Hidden, false);
            xml.WriteAttributeDefault("customBuiltin", cellStyle.BuiltInStyle is not null, true);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteDxfs(XmlTreeWriter xml, SequentialMap<int, XlsxDxfValue> differentialFormats, int lastNumFmtId)
    {
        xml.WriteStartElement("dxfs", _ns);
        xml.WriteAttribute("count", differentialFormats.Count);
        foreach (var (_, dxf) in differentialFormats.GetActual())
        {
            xml.WriteStartElement("dxf", _ns);

            if (dxf.Font != XlsxDifferentialFontValue.Empty)
                WriteFont(xml, "font", dxf.Font);

            if (dxf.NumberFormat is { } numberFormat)
            {
                // numFmtId doesn't matter in dxf, but keep them unique (Excel-like behavior)
                WriteNumFmt(xml, "numFmt", ++lastNumFmtId, numberFormat);
            }

            if (dxf.Fill != XlsxDifferentialFillValue.Empty)
                WriteFill(xml, "fill", dxf.Fill.Pattern, dxf.Fill.LinearGradient, dxf.Fill.PathGradient, true);

            if (dxf.Alignment != XlsxDifferentialAlignmentValue.Empty)
                WriteAlignment(xml, "alignment", dxf.Alignment);

            if (dxf.Border != XlsxDifferentialBorderValue.Empty)
                WriteBorder(xml, "border", dxf.Border);

            if (dxf.Protection != XlsxDifferentialProtectionValue.Empty)
                WriteProtection(xml, "protection", dxf.Protection);

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteTableStyles(XmlTreeWriter xml, SequentialMap<int, XlsxDxfValue> dxfMap, XlsxWorkbookStyles styles)
    {
        var allStyleNames = styles.TableStyles.Keys.Concat(styles.PivotStyles.Keys).Distinct(XlsxHelper.NameComparer).ToList();
        allStyleNames.Sort();

        xml.WriteStartElement("tableStyles", _ns);
        xml.WriteAttribute("count", allStyleNames.Count);
        if (styles.DefaultTableStyle is { } defaultTableStyle)
            xml.WriteAttribute("defaultTableStyle", defaultTableStyle);

        if (styles.DefaultPivotStyle is { } defaultPivotStyle)
            xml.WriteAttribute("defaultPivotStyle", defaultPivotStyle);

        foreach (var styleName in allStyleNames)
            WriteTableStyle(xml, styleName, dxfMap, styles);

        xml.WriteEndElement();
    }

    private void WriteTableStyle(XmlTreeWriter xml, string styleName, SequentialMap<int, XlsxDxfValue> dxfMap, XlsxWorkbookStyles styles)
    {
        xml.WriteStartElement("tableStyle", _ns);
        xml.WriteAttribute("name", styleName);

        var hasPivotStyle = styles.PivotStyles.TryGetValue(styleName, out var pivotStyle);
        xml.WriteAttributeDefault("pivot", hasPivotStyle, true);

        var hasTableStyle = styles.TableStyles.TryGetValue(styleName, out var tableStyle);
        xml.WriteAttributeDefault("table", hasTableStyle, true);

        var styledRegions = new List<(string Type, int Size, int DxfId)>(TableRegionsMap.Count);
        foreach (var (type, tableRegion, pivotRegion) in TableRegionsMap)
        {
            var tableRegionStyle = TryGetTableRegion(tableStyle, tableRegion);
            var pivotRegionStyle = TryGetPivotRegion(pivotStyle, pivotRegion);
            if (tableRegionStyle is var (tableSize, tableDxf) &&
                pivotRegionStyle is var (pivotSize, pivotDxf) &&
                (tableSize != pivotSize || tableDxf != pivotDxf))
            {
                // This should never happen. The table/pivot shared style have same band size and
                // dxf on load and we don't provide API to modify table/pivot styles.
                // Sidenote: Excel GUI will refuse to create table style that conflicts with a pivot
                // style and vice versa. It will show an alert 'This style name already exists.'
                throw new InvalidOperationException($"Table and pivot table style '{styleName}' that has different formatting for {tableRegion}/{pivotRegion}.");
            }

            if ((tableRegionStyle ?? pivotRegionStyle) is var (bandSize, dxf))
                styledRegions.Add((type, bandSize, dxfMap.GetSavedId(dxf)));
        }

        xml.WriteAttribute("count", styledRegions.Count);
        foreach (var (type, bandSize, dxfId) in styledRegions)
        {
            xml.WriteStartElement("tableStyleElement", _ns);
            xml.WriteAttribute("type", type);
            xml.WriteAttributeDefault("size", bandSize, 1);
            xml.WriteAttribute("dxfId", dxfId);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        static (int Size, XlsxDxfValue Dxf)? TryGetTableRegion(XlsxTableTheme? tableStyle, TableRegion? tableRegion)
        {
            if (tableStyle is null || tableRegion is null)
                return null;

            if (!tableStyle.RegionFormats.TryGetValue(tableRegion.Value, out var dxf))
                return null;

            var bandSize = tableRegion switch
            {
                TableRegion.FirstRowStripe => tableStyle.RowStripe1BandSize,
                TableRegion.SecondRowStripe => tableStyle.RowStripe2BandSize,
                TableRegion.FirstColumnStripe => tableStyle.ColumnStripe1BandSize,
                TableRegion.SecondColumnStripe => tableStyle.ColumnStripe2BandSize,
                _ => 1
            };
            return (bandSize, dxf);
        }

        static (int Size, XlsxDxfValue Dxf)? TryGetPivotRegion(XlsxPivotTableStyle? pivotStyle, PivotRegion? region)
        {
            if (pivotStyle is null || region is null)
                return null;

            if (!pivotStyle.RegionFormats.TryGetValue(region.Value, out var dxf))
                return null;

            var bandSize = region switch
            {
                PivotRegion.FirstRowStripe => pivotStyle.RowStripe1BandSize,
                PivotRegion.SecondRowStripe => pivotStyle.RowStripe2BandSize,
                PivotRegion.FirstColumnStripe => pivotStyle.ColumnStripe1BandSize,
                PivotRegion.SecondColumnStripe => pivotStyle.ColumnStripe2BandSize,
                _ => 1
            };
            return (bandSize, dxf);
        }
    }

    private void WriteColors(XmlTreeWriter xml, XlsxWorkbookStyles styles)
    {
        var hasMruColors = styles.MruColors.Count > 0;
        var indexedColors = styles.IndexedColorsArgb;
        var hasIndexColors = indexedColors is { Count: > 0 };
        if (!hasMruColors && !hasIndexColors)
            return;

        xml.WriteStartElement("colors", _ns);

        if (hasIndexColors)
        {
            xml.WriteStartElement("indexedColors", _ns);
            foreach (var indexedColor in indexedColors!)
            {
                xml.WriteStartElement("rgbColor", _ns);
                xml.WriteAttributeHex("rgb", indexedColor);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        if (hasMruColors)
        {
            xml.WriteStartElement("mruColors", _ns);
            foreach (var mruColor in styles.MruColors)
                xml.WriteColor("color", _ns, mruColor);

            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // colors
    }

    private void WriteExtensions(XmlTreeWriter xml)
    {
        xml.WriteStartElement("extLst", _ns);
        WriteSlicerStyles(xml);
        WriteTimelineStyles(xml);
        xml.WriteEndElement();
    }

    private void WriteSlicerStyles(XmlTreeWriter xml)
    {
        // TODO Styles: Represent and write back, this only writes default created by Excel
        xml.WriteStartExtension("{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}", _ns, "x14", X14Main2009SsNs);
        xml.WriteStartElement("slicerStyles", X14Main2009SsNs);
        xml.WriteAttribute("defaultSlicerStyle", "SlicerStyleLight1");
        xml.WriteEndElement();
        xml.WriteEndElement();

        // dxfs for slicer styles: 46F421CA-312F-682F-3DD2-61675219B42D
    }

    private void WriteTimelineStyles(XmlTreeWriter xml)
    {
        // TODO Styles: Represent and write back, this only writes default created by Excel
        xml.WriteStartExtension("{9260A510-F301-46a8-8635-F512D64BE5F5}", _ns, "x15", X15Main2010SsNs);
        xml.WriteStartElement("timelineStyles", X15Main2010SsNs);
        xml.WriteAttribute("defaultTimelineStyle", "TimeSlicerStyleLight1");
        xml.WriteEndElement();
        xml.WriteEndElement();

        // dxfs for timeline styles: A0A4C193-F2C1-4fcb-8827-314CF55A85BB
    }
}
