using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Columns;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.DataValidation;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Hyperlinks;
using Aid.Xlsx.Opc;
using Aid.Xlsx.PageSetup;
using Aid.Xlsx.Protection;
using Aid.Xlsx.Sparklines;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Worksheets;
using StringItem = Aid.Xlsx.CalcEngine.OneOf<string, Aid.Xlsx.RichText.XlsxImmutableRichText>;

namespace Aid.Xlsx.Internal.IO;

#nullable disable

internal static class WorksheetPartReader
{
    private static readonly XNamespace Extended = OpenXmlConst.X14Main2009SsNs;
    private static readonly XNamespace Excel = OpenXmlConst.XmMain2006;

    internal static void LoadWorksheet(XlsxWorksheet ws, Stream stream, IReadOnlyDictionary<string, Uri> hyperlinks,
        List<StringItem> sst, LoadContext context)
    {
        var fitToPage = false;
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        if (reader.MoveToContent() != XmlNodeType.Element)
            throw new InvalidDataException("The worksheet part is empty.");

        var ns = reader.NamespaceURI;
        if (reader.LocalName != "worksheet" || ns is not (OpenXmlConst.Main2006SsNs or OpenXmlConst.StrictSpreadsheetNs))
            return;
        if (reader.IsEmptyElement)
            return;

        reader.Read();
        while (!reader.EOF)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            if (reader.NamespaceURI == OpenXmlConst.MarkupCompatibilityNs &&
                reader.LocalName is "AlternateContent" or "Choice" or "Fallback")
            {
                reader.Read();
                continue;
            }

            if (reader.NamespaceURI != ns)
            {
                reader.Skip();
                continue;
            }

            switch (reader.LocalName)
            {
                case "sheetFormatPr":
                    LoadSheetFormat((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "sheetViews":
                    LoadSheetViews((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "mergeCells":
                    foreach (var merge in ((XElement)XNode.ReadFrom(reader)).Elements(XName.Get("mergeCell", ns)))
                        ws.Range((string)merge.Attribute("ref")).Merge(false);
                    break;
                case "cols":
                    LoadColumns(ws, (XElement)XNode.ReadFrom(reader));
                    break;
                case "sheetData":
                    using (var subtree = reader.ReadSubtree())
                    using (var tree = new XmlTreeReader(subtree, XmlToEnumMapper.Instance, ws.Workbook.StrictAttributeParsing, leaveOpen: true))
                        new SheetDataReader(tree, ws, ws.Workbook.Styles, sst, ns).ParseCtSheetData("sheetData", ns);
                    reader.Read();
                    break;
                case "autoFilter":
                    AutoFilterXml.LoadWorksheetFilter((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "sheetProtection":
                    LoadSheetProtection((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "dataValidations":
                    foreach (var validation in ((XElement)XNode.ReadFrom(reader)).Elements(XName.Get("dataValidation", ns)))
                        LoadDataValidation(validation, (string)validation.Attribute("sqref"), ws);
                    break;
                case "conditionalFormatting":
                    ConditionalFormattingXml.Load((XElement)XNode.ReadFrom(reader), ws, context);
                    break;
                case "hyperlinks":
                    LoadHyperlinks((XElement)XNode.ReadFrom(reader), hyperlinks, ws);
                    break;
                case "printOptions":
                    LoadPrintOptions((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "pageMargins":
                    LoadPageMargins((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "pageSetup":
                    LoadPageSetup((XElement)XNode.ReadFrom(reader), ws, fitToPage);
                    break;
                case "headerFooter":
                    LoadHeaderFooter((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "sheetPr":
                    fitToPage = LoadSheetProperties((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "rowBreaks":
                    LoadBreaks((XElement)XNode.ReadFrom(reader), ws.PageSetup.RowBreaks);
                    break;
                case "colBreaks":
                    LoadBreaks((XElement)XNode.ReadFrom(reader), ws.PageSetup.ColumnBreaks);
                    break;
                case "extLst":
                    var extensions = (XElement)XNode.ReadFrom(reader);
                    ConditionalFormattingXml.LoadDataBarExtensions(extensions, ws);
                    LoadExtensions(extensions, ws);
                    break;
                case "legacyDrawing":
                    ws.LegacyDrawingId = reader.GetAttribute("id", OpenXmlConst.RelationshipsNs)
                        ?? reader.GetAttribute("id", OpenXmlConst.StrictRelationshipsNs);
                    reader.Skip();
                    break;
                default:
                    reader.Skip();
                    break;
            }
        }
    }

    private static void LoadSheetFormat(XElement format, XlsxWorksheet ws)
    {
        if ((double?)format.Attribute("defaultRowHeight") is { } rowHeight)
            ws.RowHeight = rowHeight;

        ws.RowHeightChanged = (bool?)format.Attribute("customHeight") ?? false;

        if ((double?)format.Attribute("defaultColWidth") is { } width)
            ws.ColumnWidth = XlsxHelper.ConvertWidthToNoC(width, ws.Style.Font, ws.Workbook);
        else if ((uint?)format.Attribute("baseColWidth") is { } baseWidth)
            ws.ColumnWidth = XlsxHelper.CalculateColumnWidth(baseWidth, ws.Style.Font, ws.Workbook);
    }

    private static bool LoadSheetProperties(XElement properties, XlsxWorksheet ws)
    {
        var ns = properties.Name.Namespace;
        if (properties.Element(ns + "tabColor") is { } tabColor)
            ws.TabColor = SpreadsheetColorXml.Read(tabColor);

        var outline = properties.Element(ns + "outlinePr");
        if ((bool?)outline?.Attribute("summaryBelow") is { } below)
            ws.Outline.SummaryVLocation = below ? XlsxOutlineSummaryVLocation.Bottom : XlsxOutlineSummaryVLocation.Top;
        if ((bool?)outline?.Attribute("summaryRight") is { } right)
            ws.Outline.SummaryHLocation = right ? XlsxOutlineSummaryHLocation.Right : XlsxOutlineSummaryHLocation.Left;

        return (bool?)properties.Element(ns + "pageSetUpPr")?.Attribute("fitToPage") ?? false;
    }

    private static void LoadColumns(XlsxWorksheet ws, XElement columns)
    {
        var spans = columns.Elements(columns.Name.Namespace + "col");
        var defaultColumn = spans.FirstOrDefault(col => (uint?)col.Attribute("max") == XlsxHelper.MaxColumnNumber);

        if ((double?)defaultColumn?.Attribute("width") is { } defaultWidth)
            ws.ColumnWidth = defaultWidth - XlsxConstants.ColumnWidthOffset;

        // Sheet doesn't have a format, only column spans have format. When whole sheet is selected
        // to change format, Excel will mark all cols spans as having a particular format. Format
        // is considered a sheet format when all columns have a format and it's in the last column.
        var colSpanFormats = spans.Select(col => (MinColumn: (uint?)col.Attribute("min"), MaxColumn: (uint?)col.Attribute("max"), XfId: (uint?)col.Attribute("style") ?? 0)).ToArray();
        var allColsHaveFormat = colSpanFormats.Sum(x => x.MaxColumn - x.MinColumn + 1) == XlsxHelper.MaxColumnNumber;
        if (allColsHaveFormat)
        {
            var lastColumnXfId = colSpanFormats.Single(x => x.MaxColumn == XlsxHelper.MaxColumnNumber).XfId;
            ApplyStyle(ws, checked((int)lastColumnXfId), ws.Workbook.Styles);
        }

        foreach (var col in spans)
        {
            var max = (int)col.Attribute("max");
            if (max == XlsxHelper.MaxColumnNumber)
                continue;

            var xlColumns = (XlsxColumns)ws.Columns((int)col.Attribute("min"), max);
            xlColumns.Width = (double?)col.Attribute("width") is { } width ? width - XlsxConstants.ColumnWidthOffset : ws.ColumnWidth;
            if ((bool?)col.Attribute("hidden") == true)
                xlColumns.Hide();
            if ((bool?)col.Attribute("collapsed") == true)
                xlColumns.CollapseOnly();
            if ((int?)col.Attribute("outlineLevel") is { } outlineLevel)
                xlColumns.ForEach(c => c.OutlineLevel = outlineLevel);
            if ((uint?)col.Attribute("style") is { } styleIndex)
                ApplyStyle(xlColumns, checked((int)styleIndex), ws.Workbook.Styles);
        }
    }

    private static void LoadSheetViews(XElement views, XlsxWorksheet ws)
    {
        var ns = views.Name.Namespace;
        var view = views.Element(ns + "sheetView");
        if (view is null)
            return;

        ws.RightToLeft = (bool?)view.Attribute("rightToLeft") ?? ws.RightToLeft;
        ws.ShowFormulas = (bool?)view.Attribute("showFormulas") ?? ws.ShowFormulas;
        ws.ShowGridLines = (bool?)view.Attribute("showGridLines") ?? ws.ShowGridLines;
        ws.ShowOutlineSymbols = (bool?)view.Attribute("showOutlineSymbols") ?? ws.ShowOutlineSymbols;
        ws.ShowRowColHeaders = (bool?)view.Attribute("showRowColHeaders") ?? ws.ShowRowColHeaders;
        ws.ShowRuler = (bool?)view.Attribute("showRuler") ?? ws.ShowRuler;
        ws.ShowWhiteSpace = (bool?)view.Attribute("showWhiteSpace") ?? ws.ShowWhiteSpace;
        ws.ShowZeros = (bool?)view.Attribute("showZeros") ?? ws.ShowZeros;
        ws.TabSelected = (bool?)view.Attribute("tabSelected") ?? ws.TabSelected;

        var selection = view.Element(ns + "selection");
        if ((string)selection?.Attribute("sqref") is { } references)
            ws.Ranges(references.Replace(' ', ',')).Select();
        if ((string)selection?.Attribute("activeCell") is { } activeCell)
            ws.Cell(activeCell).SetActive();

        if ((uint?)view.Attribute("zoomScale") is { } zoom)
            ws.SheetView.ZoomScale = unchecked((int)zoom);
        if ((uint?)view.Attribute("zoomScaleNormal") is { } normalZoom)
            ws.SheetView.ZoomScaleNormal = unchecked((int)normalZoom);
        if ((uint?)view.Attribute("zoomScalePageLayoutView") is { } pageZoom)
            ws.SheetView.ZoomScalePageLayoutView = unchecked((int)pageZoom);
        if ((uint?)view.Attribute("zoomScaleSheetLayoutView") is { } sheetZoom)
            ws.SheetView.ZoomScaleSheetLayoutView = unchecked((int)sheetZoom);

        var pane = view.Element(ns + "pane");
        if ((string)pane?.Attribute("state") is "frozen" or "frozenSplit")
        {
            if ((double?)pane.Attribute("xSplit") is { } xSplit)
                ws.SheetView.SplitColumn = (int)xSplit;
            if ((double?)pane.Attribute("ySplit") is { } ySplit)
                ws.SheetView.SplitRow = (int)ySplit;
        }

        if ((string)view.Attribute("topLeftCell") is { } topLeft && XlsxHelper.IsValidA1Address(topLeft))
            ws.SheetView.TopLeftCellAddress = ws.Cell(topLeft).Address;
    }

    private static void LoadSheetProtection(XElement sp, XlsxWorksheet ws)
    {
        ws.Protection.IsProtected = (bool?)sp.Attribute("sheet") ?? false;

        var algorithmName = (string)sp.Attribute("algorithmName") ?? string.Empty;
        if (String.IsNullOrEmpty(algorithmName))
        {
            ws.Protection.PasswordHash = (string)sp.Attribute("password") ?? string.Empty;
            ws.Protection.Base64EncodedSalt = string.Empty;
        }
        else if (DescribedEnumParser<XlsxProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName))
        {
            ws.Protection.Algorithm = DescribedEnumParser<XlsxProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
            ws.Protection.PasswordHash = (string)sp.Attribute("hashValue") ?? string.Empty;
            ws.Protection.SpinCount = (uint?)sp.Attribute("spinCount") ?? 0;
            ws.Protection.Base64EncodedSalt = (string)sp.Attribute("saltValue") ?? string.Empty;
        }

        Allow("formatCells", XlsxSheetProtectionElements.FormatCells);
        Allow("formatColumns", XlsxSheetProtectionElements.FormatColumns);
        Allow("formatRows", XlsxSheetProtectionElements.FormatRows);
        Allow("insertColumns", XlsxSheetProtectionElements.InsertColumns);
        Allow("insertHyperlinks", XlsxSheetProtectionElements.InsertHyperlinks);
        Allow("insertRows", XlsxSheetProtectionElements.InsertRows);
        Allow("deleteColumns", XlsxSheetProtectionElements.DeleteColumns);
        Allow("deleteRows", XlsxSheetProtectionElements.DeleteRows);
        Allow("autoFilter", XlsxSheetProtectionElements.AutoFilter);
        Allow("pivotTables", XlsxSheetProtectionElements.PivotTables);
        Allow("sort", XlsxSheetProtectionElements.Sort);
        Allow("scenarios", XlsxSheetProtectionElements.EditScenarios);
        Allow("objects", XlsxSheetProtectionElements.EditObjects, false);
        Allow("selectLockedCells", XlsxSheetProtectionElements.SelectLockedCells, false);
        Allow("selectUnlockedCells", XlsxSheetProtectionElements.SelectUnlockedCells, false);

        void Allow(string attribute, XlsxSheetProtectionElements element, bool defaultProtected = true)
        {
            ws.Protection.AllowElement(element, !((bool?)sp.Attribute(attribute) ?? defaultProtected));
        }
    }

    private static void LoadDataValidation(XElement validation, string references, XlsxWorksheet ws)
    {
        if (string.IsNullOrWhiteSpace(references))
            return;

        var ns = validation.Name.Namespace;
        foreach (var rangeAddress in references.Split((char[])null, StringSplitOptions.RemoveEmptyEntries))
        {
            var rule = ws.DataValidations.Create(Area.Parse(rangeAddress));
            rule.IgnoreBlanks = (bool?)validation.Attribute("allowBlank") ?? rule.IgnoreBlanks;
            if ((bool?)validation.Attribute("showDropDown") is { } hideDropdown)
                rule.InCellDropdown = !hideDropdown;
            rule.ShowErrorMessage = (bool?)validation.Attribute("showErrorMessage") ?? rule.ShowErrorMessage;
            rule.ShowInputMessage = (bool?)validation.Attribute("showInputMessage") ?? rule.ShowInputMessage;
            rule.InputTitle = (string)validation.Attribute("promptTitle") ?? rule.InputTitle;
            rule.InputMessage = (string)validation.Attribute("prompt") ?? rule.InputMessage;
            rule.ErrorTitle = (string)validation.Attribute("errorTitle") ?? rule.ErrorTitle;
            rule.ErrorMessage = (string)validation.Attribute("error") ?? rule.ErrorMessage;
            rule.ErrorStyle = (string)validation.Attribute("errorStyle") switch
            {
                "stop" => XlsxErrorStyle.Stop,
                "warning" => XlsxErrorStyle.Warning,
                "information" => XlsxErrorStyle.Information,
                _ => rule.ErrorStyle,
            };
            rule.AllowedValues = (string)validation.Attribute("type") switch
            {
                "none" => XlsxAllowedValues.AnyValue,
                "whole" => XlsxAllowedValues.WholeNumber,
                "decimal" => XlsxAllowedValues.Decimal,
                "list" => XlsxAllowedValues.List,
                "date" => XlsxAllowedValues.Date,
                "time" => XlsxAllowedValues.Time,
                "textLength" => XlsxAllowedValues.TextLength,
                "custom" => XlsxAllowedValues.Custom,
                _ => rule.AllowedValues,
            };
            rule.Operator = (string)validation.Attribute("operator") switch
            {
                "between" => XlsxOperator.Between,
                "notBetween" => XlsxOperator.NotBetween,
                "equal" => XlsxOperator.EqualTo,
                "notEqual" => XlsxOperator.NotEqualTo,
                "greaterThan" => XlsxOperator.GreaterThan,
                "lessThan" => XlsxOperator.LessThan,
                "greaterThanOrEqual" => XlsxOperator.EqualOrGreaterThan,
                "lessThanOrEqual" => XlsxOperator.EqualOrLessThan,
                _ => rule.Operator,
            };
            rule.MinValue = (string)validation.Element(ns + "formula1") ?? rule.MinValue;
            rule.MaxValue = (string)validation.Element(ns + "formula2") ?? rule.MaxValue;
        }
    }

    private static void LoadHyperlinks(XElement hyperlinks, IReadOnlyDictionary<string, Uri> targets, XlsxWorksheet ws)
    {
        foreach (var link in hyperlinks.Elements(hyperlinks.Name.Namespace + "hyperlink"))
        {
            var reference = (string)link.Attribute("ref");
            if (string.IsNullOrEmpty(reference) || reference == "#REF")
                continue;

            var tooltip = (string)link.Attribute("tooltip") ?? string.Empty;
            var id = (string)link.Attribute(XName.Get("id", OpenXmlConst.RelationshipsNs))
                ?? (string)link.Attribute(XName.Get("id", OpenXmlConst.StrictRelationshipsNs));
            var location = (string)link.Attribute("location") ?? reference;
            foreach (XlsxCell cell in ws.Range(reference).Cells())
            {
                cell.SetCellHyperlink(id is null ? new XlsxHyperlink(location, tooltip) : new XlsxHyperlink(targets[id], tooltip));
            }
        }
    }

    private static void LoadPrintOptions(XElement options, XlsxWorksheet ws)
    {
        var setup = ws.PageSetup;
        setup.ShowGridlines = (bool?)options.Attribute("gridLines") ?? setup.ShowGridlines;
        setup.CenterHorizontally = (bool?)options.Attribute("horizontalCentered") ?? setup.CenterHorizontally;
        setup.CenterVertically = (bool?)options.Attribute("verticalCentered") ?? setup.CenterVertically;
        setup.ShowRowAndColumnHeadings = (bool?)options.Attribute("headings") ?? setup.ShowRowAndColumnHeadings;
    }

    private static void LoadPageMargins(XElement margins, XlsxWorksheet ws)
    {
        var target = ws.PageSetup.Margins;
        target.Bottom = (double?)margins.Attribute("bottom") ?? target.Bottom;
        target.Footer = (double?)margins.Attribute("footer") ?? target.Footer;
        target.Header = (double?)margins.Attribute("header") ?? target.Header;
        target.Left = (double?)margins.Attribute("left") ?? target.Left;
        target.Right = (double?)margins.Attribute("right") ?? target.Right;
        target.Top = (double?)margins.Attribute("top") ?? target.Top;
    }

    private static void LoadPageSetup(XElement source, XlsxWorksheet ws, bool fitToPage)
    {
        var setup = ws.PageSetup;
        if ((int?)source.Attribute("paperSize") is { } paperSize)
            setup.PaperSize = (XlsxPaperSize)paperSize;
        if ((int?)source.Attribute("scale") is { } scale)
            setup.Scale = scale;
        if (fitToPage)
        {
            setup.PagesWide = (int?)source.Attribute("fitToWidth") ?? 1;
            setup.PagesTall = (int?)source.Attribute("fitToHeight") ?? 1;
        }

        setup.PageOrder = (string)source.Attribute("pageOrder") switch
        {
            "downThenOver" => XlsxPageOrderValues.DownThenOver,
            "overThenDown" => XlsxPageOrderValues.OverThenDown,
            _ => setup.PageOrder,
        };
        setup.PageOrientation = (string)source.Attribute("orientation") switch
        {
            "default" => XlsxPageOrientation.Default,
            "portrait" => XlsxPageOrientation.Portrait,
            "landscape" => XlsxPageOrientation.Landscape,
            _ => setup.PageOrientation,
        };
        setup.ShowComments = (string)source.Attribute("cellComments") switch
        {
            "none" => XlsxShowCommentsValues.None,
            "asDisplayed" => XlsxShowCommentsValues.AsDisplayed,
            "atEnd" => XlsxShowCommentsValues.AtEnd,
            _ => setup.ShowComments,
        };
        setup.PrintErrorValue = (string)source.Attribute("errors") switch
        {
            "blank" => XlsxPrintErrorValues.Blank,
            "dash" => XlsxPrintErrorValues.Dash,
            "displayed" => XlsxPrintErrorValues.Displayed,
            "NA" => XlsxPrintErrorValues.NA,
            _ => setup.PrintErrorValue,
        };

        setup.BlackAndWhite = (bool?)source.Attribute("blackAndWhite") ?? setup.BlackAndWhite;
        setup.DraftQuality = (bool?)source.Attribute("draft") ?? setup.DraftQuality;
        if ((uint?)source.Attribute("horizontalDpi") is { } horizontalDpi)
            setup.HorizontalDpi = unchecked((int)horizontalDpi);
        if ((uint?)source.Attribute("verticalDpi") is { } verticalDpi)
            setup.VerticalDpi = unchecked((int)verticalDpi);
        if (uint.TryParse((string)source.Attribute("firstPageNumber"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var firstPageNumber))
            setup.FirstPageNumber = unchecked((int)firstPageNumber);
    }

    private static void LoadHeaderFooter(XElement headerFooter, XlsxWorksheet ws)
    {
        var setup = ws.PageSetup;
        setup.AlignHFWithMargins = (bool?)headerFooter.Attribute("alignWithMargins") ?? setup.AlignHFWithMargins;
        setup.ScaleHFWithDocument = (bool?)headerFooter.Attribute("scaleWithDoc") ?? setup.ScaleHFWithDocument;
        setup.DifferentFirstPageOnHF = (bool?)headerFooter.Attribute("differentFirst") ?? setup.DifferentFirstPageOnHF;
        setup.DifferentOddEvenPagesOnHF = (bool?)headerFooter.Attribute("differentOddEven") ?? setup.DifferentOddEvenPagesOnHF;

        LoadText((XlsxHeaderFooter)setup.Header, "Header");
        LoadText((XlsxHeaderFooter)setup.Footer, "Footer");

        void LoadText(XlsxHeaderFooter target, string suffix)
        {
            var ns = headerFooter.Name.Namespace;
            if ((string)headerFooter.Element(ns + ("odd" + suffix)) is { } odd)
                target.SetInnerText(XlsxHFOccurrence.OddPages, odd);
            if ((string)headerFooter.Element(ns + ("even" + suffix)) is { } even)
                target.SetInnerText(XlsxHFOccurrence.EvenPages, even);
            if ((string)headerFooter.Element(ns + ("first" + suffix)) is { } first)
                target.SetInnerText(XlsxHFOccurrence.FirstPage, first);
            target.SetAsInitial();
        }
    }

    private static void LoadBreaks(XElement breaks, ICollection<int> target)
    {
        foreach (var pageBreak in breaks.Elements(breaks.Name.Namespace + "brk"))
        {
            if ((int?)pageBreak.Attribute("id") is { } id)
                target.Add(id);
        }
    }

    private static void LoadExtensions(XElement extensions, XlsxWorksheet ws)
    {
        foreach (var validation in extensions.Descendants(Extended + "dataValidations").Elements(Extended + "dataValidation"))
            LoadDataValidation(validation, (string)validation.Element(Excel + "sqref"), ws);

        foreach (var source in extensions.Descendants(Extended + "sparklineGroups").Elements(Extended + "sparklineGroup"))
        {
            var group = ws.SparklineGroupsInternal.Add();
            if ((string)source.Element(Excel + "f") is { } dates)
                group.DateRange = ws.Workbook.Range(dates);

            var style = group.Style;
            style.FirstMarkerColor = ReadColor("colorFirst", style.FirstMarkerColor);
            style.LastMarkerColor = ReadColor("colorLast", style.LastMarkerColor);
            style.HighMarkerColor = ReadColor("colorHigh", style.HighMarkerColor);
            style.LowMarkerColor = ReadColor("colorLow", style.LowMarkerColor);
            style.SeriesColor = ReadColor("colorSeries", style.SeriesColor);
            style.NegativeColor = ReadColor("colorNegative", style.NegativeColor);
            style.MarkersColor = ReadColor("colorMarkers", style.MarkersColor);
            group.Style = style;

            group.DisplayHidden = (bool?)source.Attribute("displayHidden") ?? group.DisplayHidden;
            group.LineWeight = (double?)source.Attribute("lineWeight") ?? group.LineWeight;
            group.Type = (string)source.Attribute("type") switch
            {
                "line" => XlsxSparklineType.Line,
                "column" => XlsxSparklineType.Column,
                "stacked" => XlsxSparklineType.Stacked,
                _ => group.Type,
            };
            group.DisplayEmptyCellsAs = (string)source.Attribute("displayEmptyCellsAs") switch
            {
                "span" => XlsxDisplayBlanksAsValues.Interpolate,
                "gap" => XlsxDisplayBlanksAsValues.NotPlotted,
                "zero" => XlsxDisplayBlanksAsValues.Zero,
                _ => group.DisplayEmptyCellsAs,
            };

            group.ShowMarkers = ReadMarker("markers", XlsxSparklineMarkers.Markers)
                | ReadMarker("high", XlsxSparklineMarkers.HighPoint)
                | ReadMarker("low", XlsxSparklineMarkers.LowPoint)
                | ReadMarker("first", XlsxSparklineMarkers.FirstPoint)
                | ReadMarker("last", XlsxSparklineMarkers.LastPoint)
                | ReadMarker("negative", XlsxSparklineMarkers.NegativePoints);

            group.HorizontalAxis.Color = ReadColor("colorAxis", group.HorizontalAxis.Color);
            group.HorizontalAxis.IsVisible = (bool?)source.Attribute("displayXAxis") ?? group.HorizontalAxis.IsVisible;
            group.HorizontalAxis.RightToLeft = (bool?)source.Attribute("rightToLeft") ?? group.HorizontalAxis.RightToLeft;
            if ((double?)source.Attribute("manualMax") is { } manualMax)
                group.VerticalAxis.ManualMax = manualMax;
            if ((double?)source.Attribute("manualMin") is { } manualMin)
                group.VerticalAxis.ManualMin = manualMin;
            group.VerticalAxis.MinAxisType = ReadAxisType("minAxisType", group.VerticalAxis.MinAxisType);
            group.VerticalAxis.MaxAxisType = ReadAxisType("maxAxisType", group.VerticalAxis.MaxAxisType);

            foreach (var sparkline in source.Elements(Extended + "sparklines").Elements(Extended + "sparkline"))
            {
                // Excel uses the first reference when a sparkline lists multiple cells.
                var references = ((string)sparkline.Element(Excel + "sqref") ?? string.Empty).Trim().Split(' ');
                group.SetSparkline(Point.Parse(references[0]), (string)sparkline.Element(Excel + "f"));
            }

            XlsxColor ReadColor(string name, XlsxColor fallback) => source.Element(Extended + name) is { } color
                ? SpreadsheetColorXml.Read(color)
                : fallback;

            XlsxSparklineMarkers ReadMarker(string name, XlsxSparklineMarkers flag) => (bool?)source.Attribute(name) == true
                ? flag
                : XlsxSparklineMarkers.None;

            XlsxSparklineAxisMinMax ReadAxisType(string name, XlsxSparklineAxisMinMax fallback) => (string)source.Attribute(name) switch
            {
                "individual" => XlsxSparklineAxisMinMax.Automatic,
                "group" => XlsxSparklineAxisMinMax.SameForAll,
                "custom" => XlsxSparklineAxisMinMax.Custom,
                _ => fallback,
            };
        }
    }

    private static void ApplyStyle(IXlsxFormatContainer container, Int32 styleIndex, XlsxWorkbookStyles styles)
    {
        container.FormatValue = styles.CellFormats[styleIndex];
    }

    private static void ApplyStyle(XlsxColumns columns, Int32 styleIndex, XlsxWorkbookStyles styles)
    {
        // When loading columns we must propagate style to each column but not deeper. In other cases we do not propagate at all.
        foreach (XlsxColumn col in columns)
            ApplyStyle(col, styleIndex, styles);
    }
}
