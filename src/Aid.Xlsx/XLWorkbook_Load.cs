#nullable disable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Aid.Xlsx
{
    using Aid.Xlsx.Opc;
    using Drawings;

    public partial class XLWorkbook
    {
        private void Load(String file)
        {
            LoadSheets(file);
        }

        private void Load(Stream stream)
        {
            LoadSheets(stream);
        }

        private void LoadSheets(String fileName)
        {
            using var stream = File.OpenRead(fileName);
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
            using var package = new OpcSdkPackage(opc, FileAccess.Read);
            using (var dSpreadsheet = SpreadsheetDocument.Open(package))
                LoadSpreadsheetDocument(dSpreadsheet);
        }

        private void LoadSheetsFromTemplate(String fileName)
        {
            LoadSheets(fileName);
            _spreadsheetDocumentType = SpreadsheetDocumentType.Workbook;

            // If we load a workbook as a template, we have to treat it as a "new" workbook.
            // The original file will NOT be copied into place before changes are applied
            // Hence all loaded RelIds have to be cleared
            ResetAllRelIds();
        }

        private void ResetAllRelIds()
        {
            foreach (var pc in PivotCachesInternal)
                pc.WorkbookCacheRelId = null;

            var sheetId = 1u;
            foreach (var ws in WorksheetsInternal)
            {
                // Ensure unique sheetId for each sheet. 
                ws.SheetId = sheetId++;
                ws.RelId = null;

                foreach (var pt in ws.PivotTables.Cast<XLPivotTable>())
                {
                    pt.CacheDefinitionRelId = null;
                    pt.RelId = null;
                }

                foreach (var picture in ws.Pictures.Cast<XLPicture>())
                    picture.RelId = null;

                foreach (var table in ws.Tables.Cast<XLTable>())
                    table.RelId = null;
            }
        }

        private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
        {
            var context = new LoadContext();
            ShapeIdManager = new XLIdManager();

            var workbookPart = dSpreadsheet.WorkbookPart;

            if (workbookPart?.ThemePart is { } themePart)
            {
                using var themeStream = themePart.GetStream(FileMode.Open, FileAccess.Read);
                ThemeXml.Read(themeStream, Theme);
            }

            XElement workbook;
            using (var workbookStream = workbookPart.GetStream(FileMode.Open, FileAccess.Read))
                workbook = WorkbookXml.ReadRoot(workbookStream);
            WorkbookXml.LoadSettings(workbook, this);

            // Styles part is optional and each workbook must have a default style
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart is not null)
            {
                using var xmlReader = CreateTreeReader(stylesPart);
                var stylesReader = new StylesReader(xmlReader, Styles);
                stylesReader.Load();
            }
            else
            {
                Styles.Initialize();
            }

            // Spec says each package must have exactly one SST part, but some packages don't have that.
            List<OneOf<string, XLImmutableRichText>> sst = [];
            var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sstPart is not null)
            {
                using var sstXmlReader = CreateTreeReader(sstPart);
                var sstReader = new SstReader(sstXmlReader, Styles);
                sst = sstReader.ParseSst();
            }

            // TODO Styles: Verify the column width is same as DefaultColumnWidth even if normal style is missing
            ColumnWidth = XLHelper.CalculateColumnWidth(8, Format.Font, this);

            // We loop through the sheets in 2 passes: first just to add the sheets and second to add all the data for the sheets.
            // We do this mainly because it skips a very costly calculation invalidation step, but it also make things more consistent,
            // e.g. when reading calculations that reference other sheets, we know that those sheets always already exist.
            // That consistency point isn't required yet but could be taken advantage of in the future.
            var ns = workbook.Name.Namespace;
            var sheets = workbook.Element(ns + "sheets")?.Elements(ns + "sheet").Select(WorkbookXml.ReadSheet).ToArray() ?? [];
            Int32 position = 0;
            foreach (var dSheet in sheets)
            {
                position++;
                var sheetName = dSheet.Name;
                var sheetId = dSheet.SheetId;

                if (string.IsNullOrEmpty(dSheet.Id))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    var emptySheet = WorksheetsInternal.Add(sheetName, position, sheetId);
                    emptySheet.Visibility = dSheet.Visibility;

                    continue;
                }

                // Although relationship to worksheet is most common, there can be other types
                // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
                // of unsupported sheets and copy them when saving. See Codeplex #6932.
                var worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;
                if (worksheetPart == null)
                {
                    UnsupportedSheets.Add(new UnsupportedSheet { SheetId = sheetId, Position = position });
                    continue;
                }

                var ws = WorksheetsInternal.Add(sheetName, position, sheetId);
                ws.RelId = dSheet.Id;

                ws.Visibility = dSheet.Visibility;
            }

            position = 0;
            foreach (var dSheet in sheets)
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
                var worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;
                if (worksheetPart == null)
                {
                    continue;
                }

                if (!WorksheetsInternal.TryGetWorksheet(sheetName, out var ws))
                {
                    // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                    continue;
                }

                var worksheetPartReader = new WorksheetPartReader();
                worksheetPartReader.LoadWorksheet(ws, worksheetPart, sst, context);

                ws.ConditionalFormats.ReorderAccordingToOriginalPriority();

                #region LoadTables

                foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts)
                {
                    using var tableStream = tableDefinitionPart.GetStream(FileMode.Open, FileAccess.Read);
                    TableDefinitionXml.Load(tableStream, ws, worksheetPart.GetIdOfPart(tableDefinitionPart));
                }

                #endregion LoadTables

                LoadDrawings(worksheetPart, ws);

                #region LoadComments

                if (worksheetPart.WorksheetCommentsPart is { } commentsPart)
                {
                    var shapes = GetCommentShapes(worksheetPart);
                    using var commentsStream = commentsPart.GetStream(FileMode.Open, FileAccess.Read);
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
                        UnsupportedSheets.FirstOrDefault(us => us.Position == (Int32)(activeTab + 1));
                    if (unsupportedSheet != null)
                        unsupportedSheet.IsActive = true;
                    else
                    {
                        Worksheet((Int32)(activeTab + 1)).SetTabActive();
                    }
                }
            }
            LoadDefinedNames(workbook);

            // Read cache definition before table definition
            foreach (var pivotTableCacheDefinitionPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
            {
                var pivotCache = PivotTableCacheDefinitionPartReader.Load(workbookPart, pivotTableCacheDefinitionPart, this);
                if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart is { } recordsPart)
                {
                    using var reader = CreateTreeReader(recordsPart);
                    var recordsReader = new PivotCacheRecordsReader(reader, pivotCache);
                    recordsReader.ReadRecordsToCache();
                }
            }

            // Delay loading of pivot tables until all sheets have been loaded
            foreach (var dSheet in sheets)
            {
                if (string.IsNullOrEmpty(dSheet.Id))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    continue;
                }

                // The referenced sheet can also be ChartsheetPart. Only look for pivot tables in normal sheet parts.
                var worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;

                if (worksheetPart is not null)
                {
                    var ws = (XLWorksheet)WorksheetsInternal.Worksheet(dSheet.Name);

                    foreach (var pivotTablePart in worksheetPart.PivotTableParts)
                    {
                        PivotTableDefinitionPartReader.Load(workbookPart, pivotTablePart, worksheetPart, ws, context);
                    }
                }
            }
        }

        private void LoadDrawings(WorksheetPart wsPart, XLWorksheet ws)
        {
            if (wsPart.DrawingsPart != null)
            {
                var drawingsPart = wsPart.DrawingsPart;

                foreach (var anchor in drawingsPart.WorksheetDrawing.ChildElements)
                {
                    var imgId = GetImageRelIdFromAnchor(anchor);

                    //If imgId is null, we're probably dealing with a TextBox (or another shape) instead of a picture
                    if (imgId == null) continue;

                    var imagePart = drawingsPart.GetPartById(imgId);
                    using (var stream = imagePart.GetStream())
                    using (var ms = new MemoryStream())
                    {
                        stream.CopyTo(ms);
                        var vsdp = GetPropertiesFromAnchor(anchor);

                        var picture = ws.AddPicture(ms, vsdp.Name, Convert.ToInt32(vsdp.Id.Value)) as XLPicture;
                        picture.RelId = imgId;

                        Xdr.ShapeProperties spPr = anchor.Descendants<Xdr.ShapeProperties>().First();
                        picture.Placement = XLPicturePlacement.FreeFloating;

                        if (spPr?.Transform2D?.Extents?.Cx.HasValue ?? false)
                            picture.Width = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cx, ws.Workbook.DpiX);

                        if (spPr?.Transform2D?.Extents?.Cy.HasValue ?? false)
                            picture.Height = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cy, ws.Workbook.DpiY);

                        if (anchor is Xdr.AbsoluteAnchor)
                        {
                            var absoluteAnchor = anchor as Xdr.AbsoluteAnchor;
                            picture.MoveTo(
                                ConvertFromEnglishMetricUnits(absoluteAnchor.Position.X.Value, ws.Workbook.DpiX),
                                ConvertFromEnglishMetricUnits(absoluteAnchor.Position.Y.Value, ws.Workbook.DpiY)
                            );
                        }
                        else if (anchor is Xdr.OneCellAnchor)
                        {
                            var oneCellAnchor = anchor as Xdr.OneCellAnchor;
                            var from = LoadMarker(ws, oneCellAnchor.FromMarker);
                            picture.MoveTo(from.Cell, from.Offset);
                        }
                        else if (anchor is Xdr.TwoCellAnchor)
                        {
                            var twoCellAnchor = anchor as Xdr.TwoCellAnchor;
                            var from = LoadMarker(ws, twoCellAnchor.FromMarker);
                            var to = LoadMarker(ws, twoCellAnchor.ToMarker);

                            if (twoCellAnchor.EditAs == null || !twoCellAnchor.EditAs.HasValue || twoCellAnchor.EditAs.Value == Xdr.EditAsValues.TwoCell)
                            {
                                picture.MoveTo(from.Cell, from.Offset, to.Cell, to.Offset);
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)
                            {
                                var shapeProperties = twoCellAnchor.Descendants<Xdr.ShapeProperties>().FirstOrDefault();
                                if (shapeProperties != null)
                                {
                                    picture.MoveTo(
                                        ConvertFromEnglishMetricUnits(spPr.Transform2D.Offset.X, ws.Workbook.DpiX),
                                        ConvertFromEnglishMetricUnits(spPr.Transform2D.Offset.Y, ws.Workbook.DpiY)
                                    );
                                }
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell)
                            {
                                picture.MoveTo(from.Cell, from.Offset);
                            }
                        }
                    }
                }
            }
        }

        private static Int32 ConvertFromEnglishMetricUnits(long emu, double resolution)
        {
            return Convert.ToInt32(emu * resolution / 914400);
        }

        private static XLMarker LoadMarker(XLWorksheet ws, Xdr.MarkerType marker)
        {
            var row = Math.Min(XLHelper.MaxRowNumber, Math.Max(1, Convert.ToInt32(marker.RowId.InnerText) + 1));
            var column = Math.Min(XLHelper.MaxColumnNumber, Math.Max(1, Convert.ToInt32(marker.ColumnId.InnerText) + 1));
            return new XLMarker(
                ws.Cell(row, column),
                new System.Drawing.Point(
                    ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.ColumnOffset.InnerText), ws.Workbook.DpiX),
                    ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.RowOffset.InnerText), ws.Workbook.DpiY)
                )
            );
        }

        #region Comment Helpers

        private static int? GetCommentShapeId(XElement shape)
        {
            var id = (string)shape?.Attribute("id");
            if (id?.StartsWith("_x0000_s", StringComparison.Ordinal) == true)
                id = id[8..];

            return int.TryParse(id, out var shapeId) ? shapeId : null;
        }

        private void LoadCommentShape(XLComment comment, XElement shape)
        {
            LoadShapeProperties(comment, shape);
            LoadClientData(comment, shape.Elements().First(e => e.Name.LocalName == "ClientData"));
            LoadTextBox(comment, shape.Elements().First(e => e.Name.LocalName == "textbox"));

            if ((string)shape.Attribute("alt") is { } alternateText)
                comment.Style.Web.SetAlternateText(alternateText);

            LoadColorsAndLines(comment, shape);
        }

        private static IList<XElement> GetCommentShapes(WorksheetPart worksheetPart)
        {
            // Cannot get this to return Vml.Shape elements
            foreach (var vmlPart in worksheetPart.VmlDrawingParts)
            {
                using (var stream = vmlPart.GetStream(FileMode.Open))
                {
                    var xdoc = XDocumentExtensions.Load(stream);
                    if (xdoc == null)
                        continue;

                    var root = xdoc.Root.Element("xml") ?? xdoc.Root;

                    if (root == null)
                        continue;

                    var shapes = root
                        .Elements()
                        .Where(VmlDrawingPartWriter.IsCommentShape)
                        .ToList();

                    if (shapes != null)
                        return shapes;
                }
            }

            throw new ArgumentException("Could not load comments file");
        }

        #endregion Comment Helpers

        private void LoadColorsAndLines<T>(IXLDrawing<T> drawing, XElement shape)
        {
            var strokeColor = shape.Attribute(@"strokecolor");
            if (strokeColor is not null)
                drawing.Style.ColorsAndLines.LineColor = XLColor.FromVmlColor(strokeColor.Value);

            var strokeWeight = shape.Attribute(@"strokeweight");
            if (strokeWeight != null && TryGetPtValue(strokeWeight.Value, out var lineWeight))
                drawing.Style.ColorsAndLines.LineWeight = lineWeight;

            var fillColor = shape.Attribute(@"fillcolor");
            if (fillColor is not null)
                drawing.Style.ColorsAndLines.FillColor = XLColor.FromVmlColor(fillColor.Value);

            var fill = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "fill");
            if (fill != null)
            {
                var opacity = fill.Attribute("opacity");
                if (opacity != null)
                {
                    String opacityVal = opacity.Value;
                    if (opacityVal.EndsWith("f"))
                        drawing.Style.ColorsAndLines.FillTransparency =
                            Double.Parse(opacityVal.Substring(0, opacityVal.Length - 1), CultureInfo.InvariantCulture) / 65536.0;
                    else
                        drawing.Style.ColorsAndLines.FillTransparency = Double.Parse(opacityVal, CultureInfo.InvariantCulture);
                }
            }

            var stroke = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "stroke");
            if (stroke != null)
            {
                var opacity = stroke.Attribute("opacity");
                if (opacity != null)
                {
                    String opacityVal = opacity.Value;
                    if (opacityVal.EndsWith("f"))
                        drawing.Style.ColorsAndLines.LineTransparency =
                            Double.Parse(opacityVal.Substring(0, opacityVal.Length - 1), CultureInfo.InvariantCulture) / 65536.0;
                    else
                        drawing.Style.ColorsAndLines.LineTransparency = Double.Parse(opacityVal, CultureInfo.InvariantCulture);
                }

                var dashStyle = stroke.Attribute("dashstyle");
                if (dashStyle != null)
                {
                    String dashStyleVal = dashStyle.Value.ToLower();
                    if (dashStyleVal == "1 1" || dashStyleVal == "shortdot")
                    {
                        var endCap = stroke.Attribute("endcap");
                        if (endCap != null && endCap.Value == "round")
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.RoundDot;
                        else
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.SquareDot;
                    }
                    else
                    {
                        switch (dashStyleVal)
                        {
                            case "dash": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.Dash; break;
                            case "dashdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.DashDot; break;
                            case "longdash": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDash; break;
                            case "longdashdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDot; break;
                            case "longdashdotdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDotDot; break;
                        }
                    }
                }

                var lineStyle = stroke.Attribute("linestyle");
                if (lineStyle != null)
                {
                    String lineStyleVal = lineStyle.Value.ToLower();
                    switch (lineStyleVal)
                    {
                        case "single": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.Single; break;
                        case "thickbetweenthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickBetweenThin; break;
                        case "thickthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickThin; break;
                        case "thinthick": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThick; break;
                        case "thinthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThin; break;
                    }
                }
            }
        }

        private void LoadTextBox<T>(IXLDrawing<T> xlDrawing, XElement textBox)
        {
            var attStyle = textBox.Attribute("style");
            if (attStyle != null) LoadTextBoxStyle(xlDrawing, attStyle);

            var attInset = textBox.Attribute("inset");
            if (attInset != null) LoadTextBoxInset(xlDrawing, attInset);
        }

        private void LoadTextBoxInset<T>(IXLDrawing<T> xlDrawing, XAttribute attInset)
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
            foreach (var (unitName, conversion) in VmlLengthUnits)
            {
                if (unit.EndsWith(unitName) && Double.TryParse(unit[..^unitName.Length], NumberStyles.Float, CultureInfo.InvariantCulture, out var unitValue))
                {
                    var insetEmu = conversion(unitValue, dpi) ?? Emu.ZeroPt;
                    return insetEmu.To(AbsLengthUnit.Inch);
                }
            }

            // Excel treats no/unexpected unit as 0
            return 0;
        }

        private static void LoadTextBoxStyle<T>(IXLDrawing<T> xlDrawing, XAttribute attStyle)
        {
            var style = attStyle.Value;
            var attributes = style.Split(';');
            foreach (String pair in attributes)
            {
                var split = pair.Split(':');
                if (split.Length != 2) continue;

                var attribute = split[0].Trim().ToLower();
                var value = split[1].Trim();
                Boolean isVertical = false;
                switch (attribute)
                {
                    case "mso-fit-shape-to-text": xlDrawing.Style.Size.SetAutomaticSize(value.Equals("t")); break;
                    case "mso-layout-flow-alt":
                        if (value.Equals("bottom-to-top")) xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.BottomToTop);
                        else if (value.Equals("top-to-bottom")) xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.Vertical);
                        break;

                    case "layout-flow": isVertical = value.Equals("vertical"); break;
                    case "mso-direction-alt": if (value == "auto") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.Context; break;
                    case "direction": if (value == "RTL") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.RightToLeft; break;
                }
                if (isVertical && xlDrawing.Style.Alignment.Orientation == XLDrawingTextOrientation.LeftToRight)
                    xlDrawing.Style.Alignment.Orientation = XLDrawingTextOrientation.TopToBottom;
            }
        }

        private void LoadClientData<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var anchor = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Anchor");
            if (anchor != null) LoadClientDataAnchor<T>(drawing, anchor);

            LoadDrawingPositioning<T>(drawing, clientData);
            LoadDrawingProtection<T>(drawing, clientData);

            var visible = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Visible");
            drawing.Visible = visible != null &&
                              (string.IsNullOrEmpty(visible.Value) ||
                               visible.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase));

            LoadDrawingHAlignment<T>(drawing, clientData);
            LoadDrawingVAlignment<T>(drawing, clientData);
        }

        private void LoadDrawingHAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var textHAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextHAlign");
            if (textHAlign != null)
                drawing.Style.Alignment.Horizontal = (XLDrawingHorizontalAlignment)Enum.Parse(typeof(XLDrawingHorizontalAlignment), textHAlign.Value.ToProper());
        }

        private void LoadDrawingVAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var textVAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextVAlign");
            if (textVAlign != null)
                drawing.Style.Alignment.Vertical = (XLDrawingVerticalAlignment)Enum.Parse(typeof(XLDrawingVerticalAlignment), textVAlign.Value.ToProper());
        }

        private void LoadDrawingProtection<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var lockedElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Locked");
            var lockTextElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "LockText");
            Boolean locked = lockedElement != null && lockedElement.Value.ToLower() == "true";
            Boolean lockText = lockTextElement != null && lockTextElement.Value.ToLower() == "true";
            drawing.Style.Protection.Locked = locked;
            drawing.Style.Protection.LockText = lockText;
        }

        private static void LoadDrawingPositioning<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var moveWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "MoveWithCells");
            var sizeWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "SizeWithCells");
            Boolean moveWithCells = !(moveWithCellsElement != null && moveWithCellsElement.Value.ToLower() == "true");
            Boolean sizeWithCells = !(sizeWithCellsElement != null && sizeWithCellsElement.Value.ToLower() == "true");
            if (moveWithCells && !sizeWithCells)
                drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveWithCells;
            else if (moveWithCells && sizeWithCells)
                drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveAndSizeWithCells;
            else
                drawing.Style.Properties.Positioning = XLDrawingAnchor.Absolute;
        }

        private static void LoadClientDataAnchor<T>(IXLDrawing<T> drawing, XElement anchor)
        {
            var location = anchor.Value.Split(',');
            drawing.Position.Column = int.Parse(location[0]) + 1;
            drawing.Position.ColumnOffset = Double.Parse(location[1], CultureInfo.InvariantCulture) / 7.5;
            drawing.Position.Row = int.Parse(location[2]) + 1;
            drawing.Position.RowOffset = Double.Parse(location[3], CultureInfo.InvariantCulture);
        }

        private void LoadShapeProperties<T>(IXLDrawing<T> xlDrawing, XElement shape)
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
                        if (Int32.TryParse(value, out var zOrder))
                        {
                            xlDrawing.ZOrder = zOrder;
                        }
                        break;
                }
            }
        }

        private readonly Dictionary<string, double> knownUnits = new Dictionary<string, double>
        {
            {"pt", 1.0},
            {"in", 72.0},
            {"mm", 72.0/25.4}
        };

        private bool TryGetPtValue(string value, out double result)
        {
            var knownUnit = knownUnits.FirstOrDefault(ku => value.Contains(ku.Key));

            if (knownUnit.Key == null)
                return Double.TryParse(value, out result);

            value = value.Replace(knownUnit.Key, String.Empty);

            if (Double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                result *= knownUnit.Value;
                return true;
            }

            result = 0d;
            return false;
        }

        private void LoadDefinedNames(XElement workbook)
        {
            var ns = workbook.Name.Namespace;
            var definedNames = workbook.Element(ns + "definedNames");
            if (definedNames is null)
                return;

            foreach (var definedName in definedNames.Elements(ns + "definedName"))
            {
                var name = (string)definedName.Attribute("name");
                var visible = !((bool?)definedName.Attribute("hidden") ?? false);
                var localSheet = (uint?)definedName.Attribute("localSheetId");
                var localSheetId = localSheet.HasValue ? Convert.ToInt32(localSheet.Value) : -1;
                var text = definedName.Value;

                if (name == "_xlnm.Print_Area")
                {
                    var fixedNames = validateDefinedNames(text.Split(','));
                    foreach (string area in fixedNames)
                    {
                        if (area.Contains("["))
                        {
                            var ws = WorksheetsInternal.FirstOrDefault<XLWorksheet>(w => w.SheetId == (localSheetId + 1));
                            if (ws != null)
                            {
                                ws.PageSetup.PrintAreas.Add(area);
                            }
                        }
                        else
                        {
                            ParseReference(area, out String sheetName, out String sheetArea);
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
                        if (DefinedNamesInternal.All<XLDefinedName>(nr => nr.Name != name))
                            DefinedNamesInternal.Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                    }
                    else
                    {
                        if (Worksheet(localSheetId + 1).DefinedNames.All(nr => nr.Name != name))
                            ((XLDefinedNames)Worksheet(localSheetId + 1).DefinedNames).Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                    }
                }
            }
        }

        private static Regex definedNameRegex = new Regex(@"\A('?).*\1!.*\z", RegexOptions.Compiled);

        private IEnumerable<String> validateDefinedNames(IEnumerable<String> definedNames)
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
            var areas = validateDefinedNames(definition.Split(','));
            foreach (var item in areas)
            {
                if (this.Range(item) != null)
                    SetColumnsOrRowsToRepeat(item);
            }
        }

        private void SetColumnsOrRowsToRepeat(string area)
        {
            ParseReference(area, out String sheetName, out String sheetArea);
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
            if (sections.Count() == 1)
            {
                sheetName = string.Empty;
                sheetArea = item;
            }
            else
            {
                sheetName = string.Join("!", sections.Take(sections.Length - 1)).UnescapeSheetName();
                sheetArea = sections[sections.Length - 1];
            }
        }


        private XmlTreeReader CreateTreeReader(OpenXmlPart openXmlPart)
        {
            var stream = openXmlPart.GetStream(FileMode.Open);
            return new XmlTreeReader(stream, XmlToEnumMapper.Instance, StrictAttributeParsing);
        }
    }
}
