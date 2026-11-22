using System.Globalization;
using System.Xml.Linq;
using System.Xml;
using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Columns;
using Aid.Xlsx.ConditionalFormats;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Hyperlinks;
using Aid.Xlsx.Opc;
using Aid.Xlsx.PageSetup;
using Aid.Xlsx.Protection;
using Aid.Xlsx.Sparklines;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Worksheets;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Formula = DocumentFormat.OpenXml.Spreadsheet.Formula;
using SpreadsheetColumns = DocumentFormat.OpenXml.Spreadsheet.Columns;
using SpreadsheetDataValidation = DocumentFormat.OpenXml.Spreadsheet.DataValidation;
using SpreadsheetHyperlinks = DocumentFormat.OpenXml.Spreadsheet.Hyperlinks;
using StringItem = Aid.Xlsx.CalcEngine.OneOf<string, Aid.Xlsx.RichText.XlsxImmutableRichText>;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Aid.Xlsx.Internal.IO;

#nullable disable

internal class WorksheetPartReader
{
    internal void LoadWorksheet(XlsxWorksheet ws, WorksheetPart worksheetPart, List<StringItem> sst, LoadContext context)
    {
        PageSetupProperties pageSetupProperties = null;
        using var stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
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

            // Keep the cell stream independent of the DOM used by the remaining worksheet metadata handlers.
            switch (reader.LocalName)
            {
                case "sheetFormatPr":
                    LoadSheetFormat((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "sheetViews":
                    LoadSheetViews(new SheetViews(reader.ReadOuterXml()), ws);
                    break;
                case "mergeCells":
                    foreach (var cell in new MergeCells(reader.ReadOuterXml()).Elements<MergeCell>())
                        ws.Range(cell.Reference).Merge(false);
                    break;
                case "cols":
                    LoadColumns(ws, new SpreadsheetColumns(reader.ReadOuterXml()));
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
                    LoadDataValidations(new DataValidations(reader.ReadOuterXml()), ws);
                    break;
                case "conditionalFormatting":
                    LoadConditionalFormatting(new ConditionalFormatting(reader.ReadOuterXml()), ws, context);
                    break;
                case "hyperlinks":
                    LoadHyperlinks(new SpreadsheetHyperlinks(reader.ReadOuterXml()), worksheetPart, ws);
                    break;
                case "printOptions":
                    LoadPrintOptions((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "pageMargins":
                    LoadPageMargins((XElement)XNode.ReadFrom(reader), ws);
                    break;
                case "pageSetup":
                    LoadPageSetup((XElement)XNode.ReadFrom(reader), ws, pageSetupProperties);
                    break;
                case "headerFooter":
                    LoadHeaderFooter(new HeaderFooter(reader.ReadOuterXml()), ws);
                    break;
                case "sheetPr":
                    LoadSheetProperties(new SheetProperties(reader.ReadOuterXml()), ws, out pageSetupProperties);
                    break;
                case "rowBreaks":
                    LoadRowBreaks(new RowBreaks(reader.ReadOuterXml()), ws);
                    break;
                case "colBreaks":
                    LoadColumnBreaks(new ColumnBreaks(reader.ReadOuterXml()), ws);
                    break;
                case "extLst":
                    LoadExtensions(new WorksheetExtensionList(reader.ReadOuterXml()), ws);
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

    private static void LoadSheetProperties(SheetProperties sheetProperty, XlsxWorksheet ws, out PageSetupProperties pageSetupProperties)
    {
        pageSetupProperties = null;
        if (sheetProperty == null) return;

        if (sheetProperty.TabColor != null)
            ws.TabColor = sheetProperty.TabColor.ToXlsxColor();

        if (sheetProperty.OutlineProperties != null)
        {
            if (sheetProperty.OutlineProperties.SummaryBelow != null)
            {
                ws.Outline.SummaryVLocation = sheetProperty.OutlineProperties.SummaryBelow
                    ? XlsxOutlineSummaryVLocation.Bottom
                    : XlsxOutlineSummaryVLocation.Top;
            }

            if (sheetProperty.OutlineProperties.SummaryRight != null)
            {
                ws.Outline.SummaryHLocation = sheetProperty.OutlineProperties.SummaryRight
                    ? XlsxOutlineSummaryHLocation.Right
                    : XlsxOutlineSummaryHLocation.Left;
            }
        }

        if (sheetProperty.PageSetupProperties != null)
            pageSetupProperties = sheetProperty.PageSetupProperties;
    }

    private static void LoadColumns(XlsxWorksheet ws, SpreadsheetColumns columns)
    {
        if (columns == null) return;

        var wsDefaultColumn =
            columns.Elements<Column>().FirstOrDefault(c => c.Max == XlsxHelper.MaxColumnNumber);

        if (wsDefaultColumn != null && wsDefaultColumn.Width != null)
            ws.ColumnWidth = wsDefaultColumn.Width - XlsxConstants.ColumnWidthOffset;

        // Sheet doesn't have a format, only column spans have format. When whole sheet is selected
        // to change format, Excel will mark all cols spans as having a particular format. Format
        // is considered a sheet format when all columns have a format and it's in the last column.
        var colSpanFormats = columns.Elements<Column>().Select(c => (MinColumn: c.Min?.Value, MaxColumn: c.Max?.Value, XfId: c.Style?.Value ?? 0)).ToArray();
        var allColsHaveFormat = colSpanFormats.Sum(x => x.MaxColumn - x.MinColumn + 1) == XlsxHelper.MaxColumnNumber;
        if (allColsHaveFormat)
        {
            var lastColumnXfId = colSpanFormats.Single(x => x.MaxColumn == XlsxHelper.MaxColumnNumber).XfId;
            ApplyStyle(ws, checked((int)lastColumnXfId), ws.Workbook.Styles);
        }

        foreach (Column col in columns.Elements<Column>())
        {
            //IXlsxStylized toApply;
            if (col.Max == XlsxHelper.MaxColumnNumber) continue;

            var xlColumns = (XlsxColumns)ws.Columns(col.Min, col.Max);
            if (col.Width != null)
            {
                Double width = col.Width - XlsxConstants.ColumnWidthOffset;
                //if (width < 0) width = 0;
                xlColumns.Width = width;
            }
            else
                xlColumns.Width = ws.ColumnWidth;

            if (col.Hidden != null && col.Hidden)
                xlColumns.Hide();

            if (col.Collapsed != null && col.Collapsed)
                xlColumns.CollapseOnly();

            if (col.OutlineLevel != null)
            {
                var outlineLevel = col.OutlineLevel;
                xlColumns.ForEach(c => c.OutlineLevel = outlineLevel);
            }

            if (col.Style?.Value is { } styleIndex)
            {
                ApplyStyle(xlColumns, checked((int)styleIndex), ws.Workbook.Styles);
            }
        }
    }

    private static void LoadSheetViews(SheetViews sheetViews, XlsxWorksheet ws)
    {
        if (sheetViews == null) return;

        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();

        if (sheetView == null) return;

        if (sheetView.RightToLeft != null) ws.RightToLeft = sheetView.RightToLeft.Value;
        if (sheetView.ShowFormulas != null) ws.ShowFormulas = sheetView.ShowFormulas.Value;
        if (sheetView.ShowGridLines != null) ws.ShowGridLines = sheetView.ShowGridLines.Value;
        if (sheetView.ShowOutlineSymbols != null)
            ws.ShowOutlineSymbols = sheetView.ShowOutlineSymbols.Value;
        if (sheetView.ShowRowColHeaders != null) ws.ShowRowColHeaders = sheetView.ShowRowColHeaders.Value;
        if (sheetView.ShowRuler != null) ws.ShowRuler = sheetView.ShowRuler.Value;
        if (sheetView.ShowWhiteSpace != null) ws.ShowWhiteSpace = sheetView.ShowWhiteSpace.Value;
        if (sheetView.ShowZeros != null) ws.ShowZeros = sheetView.ShowZeros.Value;
        if (sheetView.TabSelected != null) ws.TabSelected = sheetView.TabSelected.Value;

        var selection = sheetView.Elements<Selection>().FirstOrDefault();
        if (selection != null)
        {
            if (selection.SequenceOfReferences != null)
                ws.Ranges(selection.SequenceOfReferences.InnerText.Replace(" ", ",")).Select();

            if (selection.ActiveCell != null)
                ws.Cell(selection.ActiveCell).SetActive();
        }

        if (sheetView.ZoomScale != null)
            ws.SheetView.ZoomScale = (int)UInt32Value.ToUInt32(sheetView.ZoomScale);
        if (sheetView.ZoomScaleNormal != null)
            ws.SheetView.ZoomScaleNormal = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleNormal);
        if (sheetView.ZoomScalePageLayoutView != null)
            ws.SheetView.ZoomScalePageLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScalePageLayoutView);
        if (sheetView.ZoomScaleSheetLayoutView != null)
            ws.SheetView.ZoomScaleSheetLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleSheetLayoutView);

        var pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (new[] { PaneStateValues.Frozen, PaneStateValues.FrozenSplit }.Contains(pane?.State?.Value ?? PaneStateValues.Split))
        {
            if (pane.HorizontalSplit != null)
                ws.SheetView.SplitColumn = (Int32)pane.HorizontalSplit.Value;
            if (pane.VerticalSplit != null)
                ws.SheetView.SplitRow = (Int32)pane.VerticalSplit.Value;
        }

        if (XlsxHelper.IsValidA1Address(sheetView.TopLeftCell))
            ws.SheetView.TopLeftCellAddress = ws.Cell(sheetView.TopLeftCell.Value).Address;
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

    /// <summary>
    /// Loads the conditional formatting.
    /// </summary>
    // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.conditionalformattingrule%28v=office.15%29.aspx?f=255&MSPPError=-2147217396
    private static void LoadConditionalFormatting(ConditionalFormatting conditionalFormatting, XlsxWorksheet ws, LoadContext context)
    {
        if (conditionalFormatting == null) return;

        var differentialFormats = ws.Workbook.Styles.DifferentialFormats;
        foreach (var fr in conditionalFormatting.Elements<ConditionalFormattingRule>())
        {
            var ranges = conditionalFormatting.SequenceOfReferences.Items
                .Select(sor => ws.Range(sor.Value));
            var conditionalFormat = new XlsxConditionalFormat(ws, XlsxAreaList.FromRanges(ws, ranges));

            conditionalFormat.StopIfTrue = OpenXmlHelper.GetBooleanValueAsBool(fr.StopIfTrue, false);

            // TODO Styles: CF with empty format is technically legal, but seriously suss. Investigate.
            if (fr.FormatId is not null)
            {
                conditionalFormat.FormatValue = differentialFormats[checked((int)fr.FormatId.Value)];
            }

            // The conditional formatting type is compulsory. If it doesn't exist, skip the entire rule.
            if (fr.Type == null) continue;
            conditionalFormat.ConditionalFormatType = fr.Type.Value.ToXlsx();
            conditionalFormat.Priority = fr.Priority?.Value ?? Int32.MaxValue;

            // Although formulas are directly used only by CellIs and Expression type, other
            // format types also write them for evaluation to the workbook, e.g. rule to
            // IsBlank writes `LEN(TRIM(A2))=0` or ContainsText writes `NOT(ISERROR(SEARCH("hello",A2)))`.
            if (conditionalFormat.ConditionalFormatType == XlsxConditionalFormatType.CellIs)
            {
                conditionalFormat.Operator = fr.Operator.Value.ToXlsx();

                // The XML schema allows up to three <formula> tags, but at most two are used.
                // Some producers emit empty <formula> tags that should be ignored and extra
                // non-empty formulas should also be ignored (Excel behavior).
                var nonEmptyFormulas = fr.Elements<Formula>()
                    .Where(static f => !String.IsNullOrEmpty(f.Text))
                    .Select<Formula, XlsxFormula>(f => GetFormula(f.Text))
                    .ToList();
                if (conditionalFormat.Operator is XlsxCFOperator.Between or XlsxCFOperator.NotBetween)
                {
                    var formulas = nonEmptyFormulas.Take(2).ToList();
                    if (formulas.Count != 2)
                        throw PartStructureException.IncorrectElementsCount();

                    conditionalFormat.Values.Add(formulas[0]);
                    conditionalFormat.Values.Add(formulas[1]);
                }
                else
                {
                    // Other XlsxCFOperators expect one argument.
                    var operatorArg = nonEmptyFormulas.FirstOrDefault();
                    if (operatorArg is null)
                        throw PartStructureException.IncorrectElementsCount();

                    conditionalFormat.Values.Add(operatorArg);
                }
            }
            else if (conditionalFormat.ConditionalFormatType == XlsxConditionalFormatType.Expression)
            {
                var formula = fr.Elements<Formula>()
                    .Where(static f => !String.IsNullOrEmpty(f.Text))
                    .FirstOrDefault();

                if (formula is null)
                    throw PartStructureException.IncorrectElementsCount();

                conditionalFormat.Values.Add(GetFormula(formula.Text));
            }
            else if (conditionalFormat.ConditionalFormatType is
                     XlsxConditionalFormatType.ContainsText or
                     XlsxConditionalFormatType.NotContainsText or
                     XlsxConditionalFormatType.StartsWith or
                     XlsxConditionalFormatType.EndsWith)
            {
                conditionalFormat.Values.Add(new XlsxFormula(fr.Text?.Value ?? string.Empty) { IsFormula = false });
            }
            else if (conditionalFormat.ConditionalFormatType == XlsxConditionalFormatType.Top10)
            {
                if (fr.Percent != null)
                    conditionalFormat.Percent = fr.Percent.Value;
                if (fr.Bottom != null)
                    conditionalFormat.Bottom = fr.Bottom.Value;
                if (fr.Rank != null)
                    conditionalFormat.Values.Add(GetFormula(fr.Rank.Value.ToString()));
            }
            else if (conditionalFormat.ConditionalFormatType == XlsxConditionalFormatType.TimePeriod)
            {
                if (fr.TimePeriod != null)
                    conditionalFormat.TimePeriod = fr.TimePeriod.Value.ToXlsx();
                else
                    conditionalFormat.TimePeriod = XlsxTimePeriod.Yesterday;
            }

            if (fr.Elements<ColorScale>().Any())
            {
                var colorScale = fr.Elements<ColorScale>().First();
                ExtractConditionalFormatValueObjects(conditionalFormat, colorScale);
            }
            else if (fr.Elements<DataBar>().Any())
            {
                var dataBar = fr.Elements<DataBar>().First();
                if (dataBar.ShowValue != null)
                    conditionalFormat.ShowBarOnly = !dataBar.ShowValue.Value;

                var id = fr.Descendants<X14.Id>().FirstOrDefault();
                if (id != null && id.Text != null && !String.IsNullOrWhiteSpace(id.Text))
                    conditionalFormat.Id = new Guid(id.Text.Substring(1, id.Text.Length - 2));

                ExtractConditionalFormatValueObjects(conditionalFormat, dataBar);
            }
            else if (fr.Elements<IconSet>().Any())
            {
                var iconSet = fr.Elements<IconSet>().First();
                if (iconSet.ShowValue != null)
                    conditionalFormat.ShowIconOnly = !iconSet.ShowValue.Value;
                if (iconSet.Reverse != null)
                    conditionalFormat.ReverseIconOrder = iconSet.Reverse.Value;

                if (iconSet.IconSetValue != null)
                    conditionalFormat.IconSetStyle = iconSet.IconSetValue.Value.ToXlsx();
                else
                    conditionalFormat.IconSetStyle = XlsxIconSetStyle.ThreeTrafficLights1;

                ExtractConditionalFormatValueObjects(conditionalFormat, iconSet);
            }

            var isPivotTableFormatting = conditionalFormatting.Pivot?.Value ?? false;
            if (isPivotTableFormatting)
                context.AddPivotTableCf(ws.Name, conditionalFormat);
            else
                ws.ConditionalFormats.Add(conditionalFormat);
        }
    }

    private static XlsxFormula GetFormula(String value)
    {
        var formula = new XlsxFormula();
        formula._value = value;
        formula.IsFormula = !(value[0] == '"' && value.EndsWith("\""));
        return formula;
    }

    private static void ExtractConditionalFormatValueObjects(XlsxConditionalFormat conditionalFormat, OpenXmlElement element)
    {
        foreach (var c in element.Elements<ConditionalFormatValueObject>())
        {
            if (c.Type != null)
                conditionalFormat.ContentTypes.Add(c.Type.Value.ToXlsx());
            if (c.Val != null)
                conditionalFormat.Values.Add(new XlsxFormula { Value = c.Val.Value });
            else
                conditionalFormat.Values.Add(null);

            if (c.GreaterThanOrEqual != null)
                conditionalFormat.IconSetOperators.Add(c.GreaterThanOrEqual.Value ? XlsxCFIconSetOperator.EqualOrGreaterThan : XlsxCFIconSetOperator.GreaterThan);
            else
                conditionalFormat.IconSetOperators.Add(XlsxCFIconSetOperator.EqualOrGreaterThan);
        }
        foreach (var c in element.Elements<Color>())
        {
            conditionalFormat.Colors.Add(c.ToXlsxColor());
        }
    }

    private static void LoadDataValidations(DataValidations dataValidations, XlsxWorksheet ws)
    {
        if (dataValidations == null) return;

        foreach (SpreadsheetDataValidation dvs in dataValidations.Elements<SpreadsheetDataValidation>())
        {
            String txt = dvs.SequenceOfReferences.InnerText;
            if (String.IsNullOrWhiteSpace(txt)) continue;
            foreach (var rangeAddress in txt.Split(' '))
            {
                var dvt = ws.DataValidations.Create(Area.Parse(rangeAddress));
                if (dvs.AllowBlank != null) dvt.IgnoreBlanks = dvs.AllowBlank;
                if (dvs.ShowDropDown != null) dvt.InCellDropdown = !dvs.ShowDropDown.Value;
                if (dvs.ShowErrorMessage != null) dvt.ShowErrorMessage = dvs.ShowErrorMessage;
                if (dvs.ShowInputMessage != null) dvt.ShowInputMessage = dvs.ShowInputMessage;
                if (dvs.PromptTitle != null) dvt.InputTitle = dvs.PromptTitle;
                if (dvs.Prompt != null) dvt.InputMessage = dvs.Prompt;
                if (dvs.ErrorTitle != null) dvt.ErrorTitle = dvs.ErrorTitle;
                if (dvs.Error != null) dvt.ErrorMessage = dvs.Error;
                if (dvs.ErrorStyle != null) dvt.ErrorStyle = dvs.ErrorStyle.Value.ToXlsx();
                if (dvs.Type != null) dvt.AllowedValues = dvs.Type.Value.ToXlsx();
                if (dvs.Operator != null) dvt.Operator = dvs.Operator.Value.ToXlsx();
                if (dvs.Formula1 != null) dvt.MinValue = dvs.Formula1.Text;
                if (dvs.Formula2 != null) dvt.MaxValue = dvs.Formula2.Text;
            }
        }
    }

    private static void LoadHyperlinks(SpreadsheetHyperlinks hyperlinks, WorksheetPart worksheetPart, XlsxWorksheet ws)
    {
        var hyperlinkDictionary = new Dictionary<String, Uri>();
        if (worksheetPart.HyperlinkRelationships != null)
            hyperlinkDictionary = worksheetPart.HyperlinkRelationships.ToDictionary(hr => hr.Id, hr => hr.Uri);

        if (hyperlinks == null) return;

        foreach (Hyperlink hl in hyperlinks.Elements<Hyperlink>())
        {
            if (hl.Reference.Value.Equals("#REF")) continue;
            String tooltip = hl.Tooltip != null ? hl.Tooltip.Value : String.Empty;
            var xlRange = ws.Range(hl.Reference.Value);
            foreach (XlsxCell xlCell in xlRange.Cells())
            {
                if (hl.Id != null)
                    xlCell.SetCellHyperlink(new XlsxHyperlink(hyperlinkDictionary[hl.Id], tooltip));
                else if (hl.Location != null)
                    xlCell.SetCellHyperlink(new XlsxHyperlink(hl.Location.Value, tooltip));
                else
                    xlCell.SetCellHyperlink(new XlsxHyperlink(hl.Reference.Value, tooltip));
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

    private static void LoadPageSetup(XElement source, XlsxWorksheet ws, PageSetupProperties properties)
    {
        var setup = ws.PageSetup;
        if ((int?)source.Attribute("paperSize") is { } paperSize)
            setup.PaperSize = (XlsxPaperSize)paperSize;
        if ((int?)source.Attribute("scale") is { } scale)
            setup.Scale = scale;
        if (properties?.FitToPage?.Value == true)
        {
            setup.PagesWide = (int?)source.Attribute("fitToWidth") ?? 1;
            setup.PagesTall = (int?)source.Attribute("fitToHeight") ?? 1;
        }

        if ((string)source.Attribute("pageOrder") is { } order)
            setup.PageOrder = new EnumValue<PageOrderValues> { InnerText = order }.Value.ToXlsx();
        if ((string)source.Attribute("orientation") is { } orientation)
            setup.PageOrientation = new EnumValue<OrientationValues> { InnerText = orientation }.Value.ToXlsx();
        if ((string)source.Attribute("cellComments") is { } comments)
            setup.ShowComments = new EnumValue<CellCommentsValues> { InnerText = comments }.Value.ToXlsx();
        if ((string)source.Attribute("errors") is { } errors)
            setup.PrintErrorValue = new EnumValue<PrintErrorValues> { InnerText = errors }.Value.ToXlsx();

        setup.BlackAndWhite = (bool?)source.Attribute("blackAndWhite") ?? setup.BlackAndWhite;
        setup.DraftQuality = (bool?)source.Attribute("draft") ?? setup.DraftQuality;
        if ((uint?)source.Attribute("horizontalDpi") is { } horizontalDpi)
            setup.HorizontalDpi = unchecked((int)horizontalDpi);
        if ((uint?)source.Attribute("verticalDpi") is { } verticalDpi)
            setup.VerticalDpi = unchecked((int)verticalDpi);
        if (uint.TryParse((string)source.Attribute("firstPageNumber"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var firstPageNumber))
            setup.FirstPageNumber = unchecked((int)firstPageNumber);
    }

    private static void LoadHeaderFooter(HeaderFooter headerFooter, XlsxWorksheet ws)
    {
        if (headerFooter == null) return;

        if (headerFooter.AlignWithMargins != null)
            ws.PageSetup.AlignHFWithMargins = headerFooter.AlignWithMargins;
        if (headerFooter.ScaleWithDoc != null)
            ws.PageSetup.ScaleHFWithDocument = headerFooter.ScaleWithDoc;

        if (headerFooter.DifferentFirst != null)
            ws.PageSetup.DifferentFirstPageOnHF = headerFooter.DifferentFirst;
        if (headerFooter.DifferentOddEven != null)
            ws.PageSetup.DifferentOddEvenPagesOnHF = headerFooter.DifferentOddEven;

        // Footers
        var xlFooter = (XlsxHeaderFooter)ws.PageSetup.Footer;
        var evenFooter = headerFooter.EvenFooter;
        if (evenFooter != null)
            xlFooter.SetInnerText(XlsxHFOccurrence.EvenPages, evenFooter.Text);
        var oddFooter = headerFooter.OddFooter;
        if (oddFooter != null)
            xlFooter.SetInnerText(XlsxHFOccurrence.OddPages, oddFooter.Text);
        var firstFooter = headerFooter.FirstFooter;
        if (firstFooter != null)
            xlFooter.SetInnerText(XlsxHFOccurrence.FirstPage, firstFooter.Text);
        // Headers
        var xlHeader = (XlsxHeaderFooter)ws.PageSetup.Header;
        var evenHeader = headerFooter.EvenHeader;
        if (evenHeader != null)
            xlHeader.SetInnerText(XlsxHFOccurrence.EvenPages, evenHeader.Text);
        var oddHeader = headerFooter.OddHeader;
        if (oddHeader != null)
            xlHeader.SetInnerText(XlsxHFOccurrence.OddPages, oddHeader.Text);
        var firstHeader = headerFooter.FirstHeader;
        if (firstHeader != null)
            xlHeader.SetInnerText(XlsxHFOccurrence.FirstPage, firstHeader.Text);

        ((XlsxHeaderFooter)ws.PageSetup.Header).SetAsInitial();
        ((XlsxHeaderFooter)ws.PageSetup.Footer).SetAsInitial();
    }

    private static void LoadRowBreaks(RowBreaks rowBreaks, XlsxWorksheet ws)
    {
        if (rowBreaks == null) return;

        foreach (Break rowBreak in rowBreaks.Elements<Break>())
            ws.PageSetup.RowBreaks.Add(Int32.Parse(rowBreak.Id.InnerText));
    }

    private static void LoadColumnBreaks(ColumnBreaks columnBreaks, XlsxWorksheet ws)
    {
        if (columnBreaks == null) return;

        foreach (Break columnBreak in columnBreaks.Elements<Break>().Where(columnBreak => columnBreak.Id != null))
        {
            ws.PageSetup.ColumnBreaks.Add(Int32.Parse(columnBreak.Id.InnerText));
        }
    }

    private static void LoadExtensions(WorksheetExtensionList extensions, XlsxWorksheet ws)
    {
        if (extensions == null)
        {
            return;
        }

        foreach (var dvs in extensions
                     .Descendants<X14.DataValidations>()
                     .SelectMany(dataValidations => dataValidations.Descendants<X14.DataValidation>()))
        {
            String txt = dvs.ReferenceSequence.InnerText;
            if (String.IsNullOrWhiteSpace(txt)) continue;
            foreach (var rangeAddress in txt.Split(' '))
            {
                var dvt = ws.DataValidations.Create(Area.Parse(rangeAddress));
                if (dvs.AllowBlank != null) dvt.IgnoreBlanks = dvs.AllowBlank;
                if (dvs.ShowDropDown != null) dvt.InCellDropdown = !dvs.ShowDropDown.Value;
                if (dvs.ShowErrorMessage != null) dvt.ShowErrorMessage = dvs.ShowErrorMessage;
                if (dvs.ShowInputMessage != null) dvt.ShowInputMessage = dvs.ShowInputMessage;
                if (dvs.PromptTitle != null) dvt.InputTitle = dvs.PromptTitle;
                if (dvs.Prompt != null) dvt.InputMessage = dvs.Prompt;
                if (dvs.ErrorTitle != null) dvt.ErrorTitle = dvs.ErrorTitle;
                if (dvs.Error != null) dvt.ErrorMessage = dvs.Error;
                if (dvs.ErrorStyle != null) dvt.ErrorStyle = dvs.ErrorStyle.Value.ToXlsx();
                if (dvs.Type != null) dvt.AllowedValues = dvs.Type.Value.ToXlsx();
                if (dvs.Operator != null) dvt.Operator = dvs.Operator.Value.ToXlsx();
                if (dvs.DataValidationForumla1 != null) dvt.MinValue = dvs.DataValidationForumla1.InnerText;
                if (dvs.DataValidationForumla2 != null) dvt.MaxValue = dvs.DataValidationForumla2.InnerText;
            }
        }

        foreach (var conditionalFormattingRule in extensions
                     .Descendants<X14.ConditionalFormattingRule>()
                     .Where(cf =>
                         cf.Type != null
                         && cf.Type.HasValue
                         && cf.Type.Value == ConditionalFormatValues.DataBar))
        {
            var xlConditionalFormat = ws.ConditionalFormats
                .Cast<XlsxConditionalFormat>()
                .SingleOrDefault(cf => cf.Id.WrapInBraces() == conditionalFormattingRule.Id);
            if (xlConditionalFormat != null)
            {
                var negativeFillColor = conditionalFormattingRule.Descendants<X14.NegativeFillColor>().SingleOrDefault();
                xlConditionalFormat.Colors.Add(negativeFillColor.ToXlsxColor());
            }
        }

        foreach (var slg in extensions
                     .Descendants<X14.SparklineGroups>()
                     .SelectMany(sparklineGroups => sparklineGroups.Descendants<X14.SparklineGroup>()))
        {
            var xlSparklineGroup = ws.SparklineGroupsInternal.Add();

            if (slg.Formula != null)
                xlSparklineGroup.DateRange = ws.Workbook.Range(slg.Formula.Text);

            var xlSparklineStyle = xlSparklineGroup.Style;
            if (slg.FirstMarkerColor != null) xlSparklineStyle.FirstMarkerColor = slg.FirstMarkerColor.ToXlsxColor();
            if (slg.LastMarkerColor != null) xlSparklineStyle.LastMarkerColor = slg.LastMarkerColor.ToXlsxColor();
            if (slg.HighMarkerColor != null) xlSparklineStyle.HighMarkerColor = slg.HighMarkerColor.ToXlsxColor();
            if (slg.LowMarkerColor != null) xlSparklineStyle.LowMarkerColor = slg.LowMarkerColor.ToXlsxColor();
            if (slg.SeriesColor != null) xlSparklineStyle.SeriesColor = slg.SeriesColor.ToXlsxColor();
            if (slg.NegativeColor != null) xlSparklineStyle.NegativeColor = slg.NegativeColor.ToXlsxColor();
            if (slg.MarkersColor != null) xlSparklineStyle.MarkersColor = slg.MarkersColor.ToXlsxColor();
            xlSparklineGroup.Style = xlSparklineStyle;

            if (slg.DisplayHidden != null) xlSparklineGroup.DisplayHidden = slg.DisplayHidden;
            if (slg.LineWeight != null) xlSparklineGroup.LineWeight = slg.LineWeight;
            if (slg.Type != null) xlSparklineGroup.Type = slg.Type.Value.ToXlsx();
            if (slg.DisplayEmptyCellsAs != null) xlSparklineGroup.DisplayEmptyCellsAs = slg.DisplayEmptyCellsAs.Value.ToXlsx();

            xlSparklineGroup.ShowMarkers = XlsxSparklineMarkers.None;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Markers, false)) xlSparklineGroup.ShowMarkers |= XlsxSparklineMarkers.Markers;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.High, false)) xlSparklineGroup.ShowMarkers |= XlsxSparklineMarkers.HighPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Low, false)) xlSparklineGroup.ShowMarkers |= XlsxSparklineMarkers.LowPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.First, false)) xlSparklineGroup.ShowMarkers |= XlsxSparklineMarkers.FirstPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Last, false)) xlSparklineGroup.ShowMarkers |= XlsxSparklineMarkers.LastPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Negative, false)) xlSparklineGroup.ShowMarkers |= XlsxSparklineMarkers.NegativePoints;

            if (slg.AxisColor != null) xlSparklineGroup.HorizontalAxis.Color = XlsxColor.FromHtml(slg.AxisColor.Rgb.Value);
            if (slg.DisplayXAxis != null) xlSparklineGroup.HorizontalAxis.IsVisible = slg.DisplayXAxis;
            if (slg.RightToLeft != null) xlSparklineGroup.HorizontalAxis.RightToLeft = slg.RightToLeft;

            if (slg.ManualMax != null) xlSparklineGroup.VerticalAxis.ManualMax = slg.ManualMax;
            if (slg.ManualMin != null) xlSparklineGroup.VerticalAxis.ManualMin = slg.ManualMin;
            if (slg.MinAxisType != null) xlSparklineGroup.VerticalAxis.MinAxisType = slg.MinAxisType.Value.ToXlsx();
            if (slg.MaxAxisType != null) xlSparklineGroup.VerticalAxis.MaxAxisType = slg.MaxAxisType.Value.ToXlsx();

            foreach (var sparkline in slg.Descendants<X14.Sparklines>().SelectMany(sparklines => sparklines.Descendants<X14.Sparkline>()))
            {
                // The sqlref must contain exactly one ref [MS-XLSX]. Excel ignores everything after the first one.
                var refText = (sparkline.ReferenceSequence?.Text ?? string.Empty).Trim().Split(' ')[0];
                var location = Point.Parse(refText);

                // Technically, there could be more than one sparkline per cell, so use Set instead of Add.
                xlSparklineGroup.SetSparkline(location, sparkline.Formula?.Text);
            }
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
