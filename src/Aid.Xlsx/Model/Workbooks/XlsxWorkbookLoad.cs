#nullable disable

using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;
using System.Xml.Linq;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Comments;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Drawings.Styles;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Opc.Packaging;
using Aid.Xlsx.Opc.Xml;
using Aid.Xlsx.SpreadsheetML.Drawings;
using Aid.Xlsx.SpreadsheetML.Pivots;
using Aid.Xlsx.SpreadsheetML.Properties;
using Aid.Xlsx.SpreadsheetML.Strings;
using Aid.Xlsx.SpreadsheetML.Styles;
using Aid.Xlsx.SpreadsheetML.Tables;
using Aid.Xlsx.SpreadsheetML.Workbooks;
using Aid.Xlsx.SpreadsheetML.Worksheets;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.Model.Workbooks;

using Drawings;

public partial class XlsxWorkbook
{
    private void Load(string file)
    {
        LoadSheets(file);
    }

    private void Load(Stream stream)
    {
        LoadSheets(stream);
    }

    private void LoadSheets(string fileName)
    {
        using FileStream stream = File.OpenRead(fileName);
        LoadSheets(stream);
    }

    private void LoadSheets(Stream stream)
    {
        if (!stream.CanSeek)
        {
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            LoadSheets(copy);
            return;
        }

        using var opc = OpcPackage.Open(stream);
        WorkbookPropertiesParts.Read(opc, this);
        LoadWorkbook(new WorkbookPackage(opc));
    }

    private void LoadSheetsFromTemplate(string fileName)
    {
        LoadSheets(fileName);
        _spreadsheetDocumentType = WorkbookPackageType.Workbook;

        // If we load a workbook as a template, we have to treat it as a "new" workbook.
        // The original file will NOT be copied into place before changes are applied
        // Hence all loaded RelIds have to be cleared
        ResetAllRelIds();
    }

    private void ResetAllRelIds()
    {
        foreach (XlsxPivotCache pc in PivotCachesInternal)
            pc.WorkbookCacheRelId = null;

        var sheetId = 1u;
        foreach (XlsxWorksheet ws in WorksheetsInternal)
        {
            // Ensure unique sheetId for each sheet.
            ws.SheetId = sheetId++;
            ws.RelId = null;

            foreach (XlsxPivotTable pt in ws.PivotTables.Cast<XlsxPivotTable>())
            {
                pt.CacheDefinitionRelId = null;
                pt.RelId = null;
            }

            foreach (XlsxPicture picture in ws.Pictures.Cast<XlsxPicture>())
                picture.RelId = null;

            foreach (XlsxTable table in ws.Tables.Cast<XlsxTable>())
                table.RelId = null;
        }
    }

    private void LoadWorkbook(WorkbookPackage package)
    {
        var context = new LoadContext();
        ShapeIdManager = new XlsxIdManager();

        OpcPart workbookPart = package.WorkbookPart ?? throw new InvalidDataException("The package has no workbook part.");

        if (package.FirstRelatedPart(workbookPart, "theme").Part is { } themePart)
        {
            using Stream themeStream = themePart.OpenRead();
            ThemeXml.Read(themeStream, Theme);
        }

        XElement workbook;
        using (Stream workbookStream = workbookPart.OpenRead())
            workbook = WorkbookXml.ReadRoot(workbookStream);
        WorkbookXml.LoadSettings(workbook, this);

        // Styles part is optional and each workbook must have a default style
        OpcPart stylesPart = package.FirstRelatedPart(workbookPart, "styles").Part;
        if (stylesPart is not null)
        {
            using XmlTreeReader xmlReader = CreateTreeReader(stylesPart);
            var stylesReader = new StylesReader(xmlReader, Styles);
            stylesReader.Load();
        }
        else
        {
            Styles.Initialize();
        }

        // Spec says each package must have exactly one SST part, but some packages don't have that.
        List<OneOf<string, XlsxImmutableRichText>> sst = [];
        OpcPart sstPart = package.FirstRelatedPart(workbookPart, "sharedStrings").Part;
        if (sstPart is not null)
        {
            using XmlTreeReader sstXmlReader = CreateTreeReader(sstPart);
            var sstReader = new SstReader(sstXmlReader, Styles);
            sst = sstReader.ParseSst();
        }

        // TODO Styles: Verify the column width is same as DefaultColumnWidth even if normal style is missing
        ColumnWidth = XlsxHelper.CalculateColumnWidth(8, Format.Font, this);

        // We loop through the sheets in 2 passes: first just to add the sheets and second to add all the data for the sheets.
        // We do this mainly because it skips a very costly calculation invalidation step, but it also make things more consistent,
        // e.g. when reading calculations that reference other sheets, we know that those sheets always already exist.
        // That consistency point isn't required yet but could be taken advantage of in the future.
        XNamespace ns = workbook.Name.Namespace;
        WorkbookXml.Sheet[] sheets = workbook.Element(ns + "sheets")?.Elements(ns + "sheet").Select(WorkbookXml.ReadSheet).ToArray() ?? [];
        int position = 0;
        foreach (WorkbookXml.Sheet dSheet in sheets)
        {
            position++;
            var sheetName = dSheet.Name;
            var sheetId = dSheet.SheetId;

            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                XlsxWorksheet emptySheet = WorksheetsInternal.Add(sheetName, position, sheetId);
                emptySheet.Visibility = dSheet.Visibility;

                continue;
            }

            // Although relationship to worksheet is most common, there can be other types
            // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
            // of unsupported sheets and copy them when saving. See Codeplex #6932.
            OpcPart worksheetPart = package.RelatedPart(workbookPart, dSheet.Id, "worksheet");
            if (worksheetPart == null)
            {
                UnsupportedSheets.Add(new UnsupportedSheet { SheetId = sheetId, Position = position });
                continue;
            }

            XlsxWorksheet ws = WorksheetsInternal.Add(sheetName, position, sheetId);
            ws.RelId = dSheet.Id;

            ws.Visibility = dSheet.Visibility;
        }

        position = 0;
        foreach (WorkbookXml.Sheet dSheet in sheets)
        {
            position++;
            var sheetName = dSheet.Name;

            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // Although relationship to worksheet is most common, there can be other types
            // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
            // of unsupported sheets and copy them when saving. See Codeplex #6932.
            OpcPart worksheetPart = package.RelatedPart(workbookPart, dSheet.Id, "worksheet");
            if (worksheetPart == null)
            {
                continue;
            }

            if (!WorksheetsInternal.TryGetWorksheet(sheetName, out XlsxWorksheet ws))
            {
                // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                continue;
            }

            using (Stream worksheetStream = worksheetPart.OpenRead())
            {
                var hyperlinks = worksheetPart.Relationships.Where(rel => WorkbookPackage.IsOfficeRelationship(rel, "hyperlink"))
                    .ToDictionary(rel => rel.Id, rel => new Uri(rel.Target, UriKind.RelativeOrAbsolute));
                WorksheetPartReader.LoadWorksheet(ws, worksheetStream, hyperlinks, sst, context);
            }

            ws.ConditionalFormats.ReorderAccordingToOriginalPriority();

            #region LoadTables

            foreach ((string id, OpcPart tableDefinitionPart) in package.RelatedParts(worksheetPart, "table"))
            {
                using Stream tableStream = tableDefinitionPart.OpenRead();
                TableDefinitionXml.Load(tableStream, ws, id);
            }

            #endregion LoadTables

            LoadDrawings(package, worksheetPart, ws);

            #region LoadComments

            if (package.FirstRelatedPart(worksheetPart, "comments").Part is { } commentsPart)
            {
                IList<XElement> shapes = GetCommentShapes(package, worksheetPart);
                using Stream commentsStream = commentsPart.OpenRead();
                CommentPartReader.Load(commentsStream, ws,
                    index => GetCommentShapeId(index < shapes.Count ? shapes[index] : null),
                    (index, comment) =>
                    {
                        if (index < shapes.Count)
                            LoadCommentShape(comment, shapes[index]);
                    });
            }

            #endregion LoadComments
        }

        if (workbook.Element(ns + "bookViews")?.Element(ns + "workbookView") is { } workbookView)
        {
            if (!uint.TryParse((string)workbookView.Attribute("activeTab"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var activeTab))
            {
                Worksheets.First().SetTabActive().Unhide();
            }
            else
            {
                UnsupportedSheet unsupportedSheet =
                    UnsupportedSheets.FirstOrDefault(us => us.Position == (int)(activeTab + 1));
                if (unsupportedSheet != null)
                    unsupportedSheet.IsActive = true;
                else
                {
                    Worksheet((int)(activeTab + 1)).SetTabActive();
                }
            }
        }
        LoadDefinedNames(workbook);

        // Read cache definition before table definition
        var cacheRelationships = package.RelatedParts(workbookPart, "pivotCacheDefinition").ToList();
        foreach ((string id, OpcPart pivotTableCacheDefinitionPart) in cacheRelationships)
        {
            using Stream definitionStream = pivotTableCacheDefinitionPart.OpenRead();
            XElement definition = PivotCacheDefinitionXml.Read(definitionStream);
            XlsxPivotCache pivotCache = PivotCacheDefinitionXml.Load(definition, id, this);
            if (package.FirstRelatedPart(pivotTableCacheDefinitionPart, "pivotCacheRecords").Part is { } recordsPart)
            {
                using XmlTreeReader reader = CreateTreeReader(recordsPart);
                var recordsReader = new PivotCacheRecordsReader(reader, pivotCache);
                recordsReader.ReadRecordsToCache();
            }
        }

        // Delay loading of pivot tables until all sheets have been loaded
        foreach (WorkbookXml.Sheet dSheet in sheets)
        {
            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // The referenced sheet can also be ChartsheetPart. Only look for pivot tables in normal sheet parts.
            OpcPart worksheetPart = package.RelatedPart(workbookPart, dSheet.Id, "worksheet");

            if (worksheetPart is not null)
            {
                var ws = (XlsxWorksheet)WorksheetsInternal.Worksheet(dSheet.Name);

                foreach ((string id, OpcPart pivotTablePart) in package.RelatedParts(worksheetPart, "pivotTable"))
                {
                    (string definitionId, OpcPart cachePart) = package.FirstRelatedPart(pivotTablePart, "pivotCacheDefinition");
                    if (cachePart is null)
                        continue;

                    var cacheId = cacheRelationships.FirstOrDefault(pair => pair.Part == cachePart).Id;
                    XlsxPivotCache cache = PivotCachesInternal.FirstOrDefault<XlsxPivotCache>(candidate => candidate.WorkbookCacheRelId == cacheId);
                    if (cache is null)
                    {
                        // A producer may use distinct caches for the same source range or table.
                        using Stream cacheStream = cachePart.OpenRead();
                        IXlsxPivotSource source = PivotCacheDefinitionXml.ReadSource(PivotCacheDefinitionXml.Read(cacheStream));
                        cache = PivotCachesInternal.FirstOrDefault<XlsxPivotCache>(candidate => candidate.Source.Equals(source));
                    }

                    if (cache is not null)
                    {
                        using Stream partStream = pivotTablePart.OpenRead();
                        PivotTableXml.Load(PivotTableXml.Read(partStream), ws, cache, id, definitionId, context);
                    }
                }
            }
        }
    }

    private static void LoadDrawings(WorkbookPackage package, OpcPart worksheetPart, XlsxWorksheet sheet)
    {
        if (package.FirstRelatedPart(worksheetPart, "drawing").Part is not { } part)
            return;

        using Stream stream = part.OpenRead();
        DrawingXml.LoadPictures(DrawingXml.Read(stream), sheet,
            id => package.RelatedPart(part, id, "image")?.OpenRead());
    }

    #region Comment Helpers

    private static int? GetCommentShapeId(XElement shape)
    {
        var id = (string)shape?.Attribute("id");
        if (id?.StartsWith("_x0000_s", StringComparison.Ordinal) == true)
            id = id[8..];

        return int.TryParse(id, out var shapeId) ? shapeId : null;
    }

    private void LoadCommentShape(XlsxComment comment, XElement shape)
    {
        LoadShapeProperties(comment, shape);
        LoadClientData(comment, shape.Elements().First(e => e.Name.LocalName == "ClientData"));
        LoadTextBox(comment, shape.Elements().First(e => e.Name.LocalName == "textbox"));

        if ((string)shape.Attribute("alt") is { } alternateText)
            comment.Style.Web.SetAlternateText(alternateText);

        LoadColorsAndLines(comment, shape);
    }

    private static IList<XElement> GetCommentShapes(WorkbookPackage package, OpcPart worksheetPart)
    {
        foreach ((string _, OpcPart vmlPart) in package.RelatedParts(worksheetPart, "vmlDrawing"))
        {
            using Stream stream = vmlPart.OpenRead();
            XDocument xdoc = XDocumentExtensions.Load(stream);
            if (xdoc == null)
                continue;

            XElement root = xdoc.Root.Element("xml") ?? xdoc.Root;

            if (root == null)
                continue;

            var shapes = root
                .Elements()
                .Where(VmlDrawingPartWriter.IsCommentShape)
                .ToList();

            if (shapes != null)
                return shapes;
        }

        throw new ArgumentException("Could not load comments file");
    }

    #endregion Comment Helpers

    private void LoadColorsAndLines<T>(IXlsxDrawing<T> drawing, XElement shape)
    {
        XAttribute strokeColor = shape.Attribute(@"strokecolor");
        if (strokeColor is not null)
            drawing.Style.ColorsAndLines.LineColor = XlsxColor.FromVmlColor(strokeColor.Value);

        XAttribute strokeWeight = shape.Attribute(@"strokeweight");
        if (strokeWeight != null && TryGetPtValue(strokeWeight.Value, out var lineWeight))
            drawing.Style.ColorsAndLines.LineWeight = lineWeight;

        XAttribute fillColor = shape.Attribute(@"fillcolor");
        if (fillColor is not null)
            drawing.Style.ColorsAndLines.FillColor = XlsxColor.FromVmlColor(fillColor.Value);

        XElement fill = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "fill");
        if (fill != null)
        {
            XAttribute opacity = fill.Attribute("opacity");
            if (opacity != null)
            {
                string opacityVal = opacity.Value;
                if (opacityVal.EndsWith('f'))
                    drawing.Style.ColorsAndLines.FillTransparency =
                        double.Parse(opacityVal[..^1], CultureInfo.InvariantCulture) / 65536.0;
                else
                    drawing.Style.ColorsAndLines.FillTransparency = double.Parse(opacityVal, CultureInfo.InvariantCulture);
            }
        }

        XElement stroke = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "stroke");
        if (stroke != null)
        {
            XAttribute opacity = stroke.Attribute("opacity");
            if (opacity != null)
            {
                string opacityVal = opacity.Value;
                if (opacityVal.EndsWith('f'))
                    drawing.Style.ColorsAndLines.LineTransparency =
                        double.Parse(opacityVal[..^1], CultureInfo.InvariantCulture) / 65536.0;
                else
                    drawing.Style.ColorsAndLines.LineTransparency = double.Parse(opacityVal, CultureInfo.InvariantCulture);
            }

            XAttribute dashStyle = stroke.Attribute("dashstyle");
            if (dashStyle != null)
            {
                string dashStyleVal = dashStyle.Value.ToLower();
                if (dashStyleVal == "1 1" || dashStyleVal == "shortdot")
                {
                    XAttribute endCap = stroke.Attribute("endcap");
                    if (endCap != null && endCap.Value == "round")
                        drawing.Style.ColorsAndLines.LineDash = XlsxDashStyle.RoundDot;
                    else
                        drawing.Style.ColorsAndLines.LineDash = XlsxDashStyle.SquareDot;
                }
                else
                {
                    switch (dashStyleVal)
                    {
                        case "dash": drawing.Style.ColorsAndLines.LineDash = XlsxDashStyle.Dash; break;
                        case "dashdot": drawing.Style.ColorsAndLines.LineDash = XlsxDashStyle.DashDot; break;
                        case "longdash": drawing.Style.ColorsAndLines.LineDash = XlsxDashStyle.LongDash; break;
                        case "longdashdot": drawing.Style.ColorsAndLines.LineDash = XlsxDashStyle.LongDashDot; break;
                        case "longdashdotdot": drawing.Style.ColorsAndLines.LineDash = XlsxDashStyle.LongDashDotDot; break;
                    }
                }
            }

            XAttribute lineStyle = stroke.Attribute("linestyle");
            if (lineStyle != null)
            {
                string lineStyleVal = lineStyle.Value.ToLower();
                switch (lineStyleVal)
                {
                    case "single": drawing.Style.ColorsAndLines.LineStyle = XlsxLineStyle.Single; break;
                    case "thickbetweenthin": drawing.Style.ColorsAndLines.LineStyle = XlsxLineStyle.ThickBetweenThin; break;
                    case "thickthin": drawing.Style.ColorsAndLines.LineStyle = XlsxLineStyle.ThickThin; break;
                    case "thinthick": drawing.Style.ColorsAndLines.LineStyle = XlsxLineStyle.ThinThick; break;
                    case "thinthin": drawing.Style.ColorsAndLines.LineStyle = XlsxLineStyle.ThinThin; break;
                }
            }
        }
    }

    private void LoadTextBox<T>(IXlsxDrawing<T> xlDrawing, XElement textBox)
    {
        XAttribute attStyle = textBox.Attribute("style");
        if (attStyle != null) LoadTextBoxStyle(xlDrawing, attStyle);

        XAttribute attInset = textBox.Attribute("inset");
        if (attInset != null) LoadTextBoxInset(xlDrawing, attInset);
    }

    private void LoadTextBoxInset<T>(IXlsxDrawing<T> xlDrawing, XAttribute attInset)
    {
        var split = attInset.Value.Split(',');
        xlDrawing.Style.Margins.Left = GetInsetInInches(split[0], DpiX);
        xlDrawing.Style.Margins.Top = GetInsetInInches(split[1], DpiY);
        xlDrawing.Style.Margins.Right = GetInsetInInches(split[2], DpiX);
        xlDrawing.Style.Margins.Bottom = GetInsetInInches(split[3], DpiY);
    }

    /// <summary>
    /// List of all VML length units and their conversion. Key is a name, value is a conversion
    /// function to EMU. See <a href="https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-units">documentation</a>.
    /// </summary>
    /// <remarks>
    /// OI-29500 says <em>Office also uses EMUs throughout VML as a valid unit system</em>.
    /// Relative units conversions are guesstimated by how Excel 2022 behaves for inset
    /// attribute of <c>TextBox</c> element of a note/comment. Generally speaking, Excel
    /// converts relative values to physical length (e.g. <c>px</c> to <c>pt</c>) and saves
    /// them as such. The <c>ex</c>/<c>em</c> units are not interpreted as described in the
    /// doc, but as 1/90th or an inch. The <c>%</c> seems to be always 0.
    /// </remarks>
    private static readonly Dictionary<string, Func<double, double, Emu?>> VmlLengthUnits = new()
    {
        {"in", (value, _) => Emu.From(value, AbsLengthUnit.Inch) },
        {"cm", (value, _) => Emu.From(value, AbsLengthUnit.Centimeter) },
        {"mm", (value, _) => Emu.From(value, AbsLengthUnit.Millimeter) },
        {"pt", (value, _) => Emu.From(value, AbsLengthUnit.Point) },
        {"pc", (value, _) => Emu.From(value, AbsLengthUnit.Pica) },
        {"emu", (value, _) => Emu.From(value , AbsLengthUnit.Emu) },
        {"px", (value, dpi) => Emu.From(value / dpi, AbsLengthUnit.Inch) },
        {"em", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
        {"ex", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
        {"%", (_, _) => Emu.ZeroPt },
    };

    private static double GetInsetInInches(string value, double dpi)
    {
        var unit = value.Trim();
        foreach ((string unitName, Func<double, double, Emu?> conversion) in VmlLengthUnits)
        {
            if (unit.EndsWith(unitName) && double.TryParse(unit[..^unitName.Length], NumberStyles.Float, CultureInfo.InvariantCulture, out var unitValue))
            {
                Emu insetEmu = conversion(unitValue, dpi) ?? Emu.ZeroPt;
                return insetEmu.To(AbsLengthUnit.Inch);
            }
        }

        // Excel treats no/unexpected unit as 0
        return 0;
    }

    private static void LoadTextBoxStyle<T>(IXlsxDrawing<T> xlDrawing, XAttribute attStyle)
    {
        var style = attStyle.Value;
        var attributes = style.Split(';');
        foreach (string pair in attributes)
        {
            var split = pair.Split(':');
            if (split.Length != 2) continue;

            var attribute = split[0].Trim().ToLower();
            var value = split[1].Trim();
            bool isVertical = false;
            switch (attribute)
            {
                case "mso-fit-shape-to-text": xlDrawing.Style.Size.SetAutomaticSize(value.Equals("t")); break;
                case "mso-layout-flow-alt":
                    if (value.Equals("bottom-to-top")) xlDrawing.Style.Alignment.SetOrientation(XlsxDrawingTextOrientation.BottomToTop);
                    else if (value.Equals("top-to-bottom")) xlDrawing.Style.Alignment.SetOrientation(XlsxDrawingTextOrientation.Vertical);
                    break;

                case "layout-flow": isVertical = value.Equals("vertical"); break;
                case "mso-direction-alt": if (value == "auto") xlDrawing.Style.Alignment.Direction = XlsxDrawingTextDirection.Context; break;
                case "direction": if (value == "RTL") xlDrawing.Style.Alignment.Direction = XlsxDrawingTextDirection.RightToLeft; break;
            }
            if (isVertical && xlDrawing.Style.Alignment.Orientation == XlsxDrawingTextOrientation.LeftToRight)
                xlDrawing.Style.Alignment.Orientation = XlsxDrawingTextOrientation.TopToBottom;
        }
    }

    private void LoadClientData<T>(IXlsxDrawing<T> drawing, XElement clientData)
    {
        XElement anchor = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Anchor");
        if (anchor != null) LoadClientDataAnchor<T>(drawing, anchor);

        LoadDrawingPositioning<T>(drawing, clientData);
        LoadDrawingProtection<T>(drawing, clientData);

        XElement visible = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Visible");
        drawing.Visible = visible != null &&
                          (string.IsNullOrEmpty(visible.Value) ||
                           visible.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase));

        LoadDrawingHAlignment<T>(drawing, clientData);
        LoadDrawingVAlignment<T>(drawing, clientData);
    }

    private void LoadDrawingHAlignment<T>(IXlsxDrawing<T> drawing, XElement clientData)
    {
        XElement textHAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextHAlign");
        if (textHAlign != null)
            drawing.Style.Alignment.Horizontal = Enum.Parse<XlsxDrawingHorizontalAlignment>(textHAlign.Value.ToProper());
    }

    private void LoadDrawingVAlignment<T>(IXlsxDrawing<T> drawing, XElement clientData)
    {
        XElement textVAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextVAlign");
        if (textVAlign != null)
            drawing.Style.Alignment.Vertical = Enum.Parse<XlsxDrawingVerticalAlignment>(textVAlign.Value.ToProper());
    }

    private void LoadDrawingProtection<T>(IXlsxDrawing<T> drawing, XElement clientData)
    {
        XElement lockedElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Locked");
        XElement lockTextElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "LockText");
        bool locked = lockedElement != null && lockedElement.Value.Equals("true", StringComparison.CurrentCultureIgnoreCase);
        bool lockText = lockTextElement != null && lockTextElement.Value.Equals("true", StringComparison.CurrentCultureIgnoreCase);
        drawing.Style.Protection.Locked = locked;
        drawing.Style.Protection.LockText = lockText;
    }

    private static void LoadDrawingPositioning<T>(IXlsxDrawing<T> drawing, XElement clientData)
    {
        XElement moveWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "MoveWithCells");
        XElement sizeWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "SizeWithCells");
        bool moveWithCells = !(moveWithCellsElement != null && moveWithCellsElement.Value.Equals("true", StringComparison.CurrentCultureIgnoreCase));
        bool sizeWithCells = !(sizeWithCellsElement != null && sizeWithCellsElement.Value.Equals("true", StringComparison.CurrentCultureIgnoreCase));
        if (moveWithCells && !sizeWithCells)
            drawing.Style.Properties.Positioning = XlsxDrawingAnchor.MoveWithCells;
        else if (moveWithCells && sizeWithCells)
            drawing.Style.Properties.Positioning = XlsxDrawingAnchor.MoveAndSizeWithCells;
        else
            drawing.Style.Properties.Positioning = XlsxDrawingAnchor.Absolute;
    }

    private static void LoadClientDataAnchor<T>(IXlsxDrawing<T> drawing, XElement anchor)
    {
        var location = anchor.Value.Split(',');
        drawing.Position.Column = int.Parse(location[0]) + 1;
        drawing.Position.ColumnOffset = double.Parse(location[1], CultureInfo.InvariantCulture) / 7.5;
        drawing.Position.Row = int.Parse(location[2]) + 1;
        drawing.Position.RowOffset = double.Parse(location[3], CultureInfo.InvariantCulture);
    }

    private void LoadShapeProperties<T>(IXlsxDrawing<T> xlDrawing, XElement shape)
    {
        if (shape.Attribute("style") == null)
            return;

        foreach (var attributePair in shape.Attribute("style").Value.Split(';'))
        {
            var split = attributePair.Split(':');
            if (split.Length != 2) continue;

            var attribute = split[0].Trim().ToLower();
            var value = split[1].Trim();

            switch (attribute)
            {
                case "visibility": xlDrawing.Visible = string.Equals("visible", value, StringComparison.OrdinalIgnoreCase); break;
                case "width":
                    if (TryGetPtValue(value, out var ptWidth))
                    {
                        xlDrawing.Style.Size.Width = ptWidth / 7.5;
                    }
                    break;

                case "height":
                    if (TryGetPtValue(value, out var ptHeight))
                    {
                        xlDrawing.Style.Size.Height = ptHeight;
                    }
                    break;

                case "z-index":
                    if (int.TryParse(value, out var zOrder))
                    {
                        xlDrawing.ZOrder = zOrder;
                    }
                    break;
            }
        }
    }

    private readonly Dictionary<string, double> knownUnits = new()
    {
        {"pt", 1.0},
        {"in", 72.0},
        {"mm", 72.0/25.4}
    };

    private bool TryGetPtValue(string value, out double result)
    {
        KeyValuePair<string, double> knownUnit = knownUnits.FirstOrDefault(ku => value.Contains(ku.Key));

        if (knownUnit.Key == null)
            return double.TryParse(value, out result);

        value = value.Replace(knownUnit.Key, string.Empty);

        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
        {
            result *= knownUnit.Value;
            return true;
        }

        result = 0d;
        return false;
    }

    private void LoadDefinedNames(XElement workbook)
    {
        XNamespace ns = workbook.Name.Namespace;
        XElement definedNames = workbook.Element(ns + "definedNames");
        if (definedNames is null)
            return;

        foreach (XElement definedName in definedNames.Elements(ns + "definedName"))
        {
            var name = (string)definedName.Attribute("name");
            var visible = !((bool?)definedName.Attribute("hidden") ?? false);
            var localSheet = (uint?)definedName.Attribute("localSheetId");
            var localSheetId = localSheet.HasValue ? Convert.ToInt32(localSheet.Value) : -1;
            var text = definedName.Value;

            if (name == "_xlnm.Print_Area")
            {
                IEnumerable<string> fixedNames = ValidateDefinedNames(text.Split(','));
                foreach (string area in fixedNames)
                {
                    if (area.Contains('['))
                    {
                        XlsxWorksheet ws = WorksheetsInternal.FirstOrDefault<XlsxWorksheet>(w => w.SheetId == (localSheetId + 1));
                        ws?.PageSetup.PrintAreas.Add(area);
                    }
                    else
                    {
                        ParseReference(area, out string sheetName, out string sheetArea);
                        if (!(sheetArea.Equals("#REF") || sheetArea.EndsWith("#REF!") || sheetArea.Length == 0 || sheetName.Length == 0))
                            WorksheetsInternal.Worksheet(sheetName).PageSetup.PrintAreas.Add(sheetArea);
                    }
                }
            }
            else if (name == "_xlnm.Print_Titles")
            {
                LoadPrintTitles(text);
            }
            else
            {
                var comment = (string)definedName.Attribute("comment");
                if (localSheetId == -1)
                {
                    if (DefinedNamesInternal.All<XlsxDefinedName>(nr => nr.Name != name))
                        DefinedNamesInternal.Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                }
                else
                {
                    if (Worksheet(localSheetId + 1).DefinedNames.All(nr => nr.Name != name))
                        ((XlsxDefinedNames)Worksheet(localSheetId + 1).DefinedNames).Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                }
            }
        }
    }

    private static readonly Regex definedNameRegex = MyRegex();

    private IEnumerable<string> ValidateDefinedNames(IEnumerable<string> definedNames)
    {
        var sb = new StringBuilder();
        foreach (string testName in definedNames)
        {
            if (sb.Length > 0)
                sb.Append(',');

            sb.Append(testName);

            Match matchedValidPattern = definedNameRegex.Match(sb.ToString());
            if (matchedValidPattern.Success)
            {
                yield return sb.ToString();
                sb = new StringBuilder();
            }
        }

        if (sb.Length > 0)
            yield return sb.ToString();
    }

    private void LoadPrintTitles(string definition)
    {
        IEnumerable<string> areas = ValidateDefinedNames(definition.Split(','));
        foreach (var item in areas)
        {
            if (this.Range(item) != null)
                SetColumnsOrRowsToRepeat(item);
        }
    }

    private void SetColumnsOrRowsToRepeat(string area)
    {
        ParseReference(area, out string sheetName, out string sheetArea);
        sheetArea = sheetArea.Replace("$", "");

        if (sheetArea.Equals("#REF")) return;
        if (IsColReference(sheetArea))
            WorksheetsInternal.Worksheet(sheetName).PageSetup.SetColumnsToRepeatAtLeft(sheetArea);
        if (IsRowReference(sheetArea))
            WorksheetsInternal.Worksheet(sheetName).PageSetup.SetRowsToRepeatAtTop(sheetArea);
    }

    // either $A:$X => true or $1:$99 => false
    private static bool IsColReference(string sheetArea)
    {
        return sheetArea.All(c => c == ':' || char.IsLetter(c));
    }

    private static bool IsRowReference(string sheetArea)
    {
        return sheetArea.All(c => c == ':' || char.IsNumber(c));
    }

    private static void ParseReference(string item, out string sheetName, out string sheetArea)
    {
        var sections = item.Trim().Split('!');
        if (sections.Length == 1)
        {
            sheetName = string.Empty;
            sheetArea = item;
        }
        else
        {
            sheetName = string.Join("!", sections.Take(sections.Length - 1)).UnescapeSheetName();
            sheetArea = sections[^1];
        }
    }


    private XmlTreeReader CreateTreeReader(OpcPart part)
    {
        Stream stream = part.OpenRead();
        return new XmlTreeReader(stream, XmlToEnumMapper.Instance, StrictAttributeParsing);
    }

    [GeneratedRegex(@"\A('?).*\1!.*\z", RegexOptions.Compiled)]
    private static partial Regex MyRegex();
}
