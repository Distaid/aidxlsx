using System.Xml;
using System.Xml.Linq;
using static Aid.Xlsx.SpreadsheetML.Xml.OpenXmlConst;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.SpreadsheetML.Styles;
using Aid.Xlsx.SpreadsheetML.Tables;

namespace Aid.Xlsx.SpreadsheetML.Worksheets;

internal sealed partial class WorksheetXml(XElement root)
{
    private static readonly XNamespace Main = Main2006SsNs;
    private static readonly XNamespace Relationships = RelationshipsNs;
    private static readonly XNamespace Extended = X14Main2009SsNs;
    private static readonly XNamespace Excel = XmMain2006;
    private static readonly XNamespace Compatibility = MarkupCompatibilityNs;

    private static readonly string[] ChildOrder =
    [
        "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
        "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
        "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
        "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
        "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
        "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "AlternateContent",
        "webPublishItems", "tableParts", "extLst",
    ];

    internal XElement Root { get; } = root;

    internal static WorksheetXml Create(Stream? source, bool writesRowExtension)
    {
        XElement root;
        if (source is null)
        {
            root = new XElement(Main + "worksheet", new XAttribute(XNamespace.Xmlns + "x", Main));
        }
        else
        {
            using var reader = XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
            reader.MoveToContent();
            var ns = reader.NamespaceURI;
            if (reader.LocalName != "worksheet" || ns is not (Main2006SsNs or StrictSpreadsheetNs))
                throw new InvalidDataException("Expected a SpreadsheetML worksheet part.");

            root = new XElement(XName.Get(reader.LocalName, ns));
            if (reader.MoveToFirstAttribute())
            {
                do
                {
                    XName name = reader.Name == "xmlns" ? XName.Get("xmlns") : XName.Get(reader.LocalName, reader.NamespaceURI);
                    root.Add(new XAttribute(name, reader.Value));
                } while (reader.MoveToNextAttribute());
                reader.MoveToElement();
            }

            // Only keep the cell data position. The cells are streamed from the model on save.
            if (reader.Read() && reader.Depth > 0)
            {
                while (!reader.EOF && reader.Depth > 0)
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetData" && reader.NamespaceURI == ns)
                    {
                        root.Add(new XElement(XName.Get("sheetData", ns)));
                        reader.Skip();
                    }
                    else if (reader.NodeType == XmlNodeType.Element)
                    {
                        using (XmlReader subtree = reader.ReadSubtree())
                            root.Add(XElement.Load(subtree, LoadOptions.PreserveWhitespace));
                        reader.Read();
                    }
                    else if (reader.NodeType is XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace)
                    {
                        root.Add(new XText(reader.Value));
                        reader.Read();
                    }
                    else
                    {
                        root.Add(XNode.ReadFrom(reader));
                    }
                }
            }

            if (ns == StrictSpreadsheetNs)
                NormalizeStrictNamespaces(root);
        }

        EnsurePrefix(root, Relationships, "r");
        EnsurePrefix(root, Compatibility, "mc");
        var hadRowExtensionPrefix = root.GetPrefixOfNamespace(X14Ac2009SsNs) is { Length: > 0 };
        var rowExtensionPrefix = EnsurePrefix(root, X14Ac2009SsNs, "x14ac");
        var ignorable = (string?)root.Attribute(Compatibility + "Ignorable");
        if ((!hadRowExtensionPrefix || writesRowExtension)
            && ignorable?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Contains(rowExtensionPrefix) != true)
            root.SetAttributeValue(Compatibility + "Ignorable", string.IsNullOrWhiteSpace(ignorable) ? rowExtensionPrefix : ignorable + " " + rowExtensionPrefix);

        return new WorksheetXml(root);
    }

    private static string EnsurePrefix(XElement element, XNamespace ns, string preferred)
    {
        if (element.GetPrefixOfNamespace(ns) is { Length: > 0 } existing)
            return existing;

        var prefix = preferred;
        for (var suffix = 1; element.GetNamespaceOfPrefix(prefix) is not null; suffix++)
            prefix = preferred + suffix.ToInvariantString();
        element.Add(new XAttribute(XNamespace.Xmlns + prefix, ns.NamespaceName));
        return prefix;
    }

    private static void NormalizeStrictNamespaces(XElement root)
    {
        foreach (XElement element in root.DescendantsAndSelf())
        {
            if (element.Name.NamespaceName == StrictSpreadsheetNs)
                element.Name = Main + element.Name.LocalName;
            foreach (XAttribute? attribute in element.Attributes().ToArray())
            {
                if (attribute.IsNamespaceDeclaration)
                {
                    if (attribute.Value == StrictSpreadsheetNs)
                        attribute.Value = Main2006SsNs;
                    else if (attribute.Value == StrictRelationshipsNs)
                        attribute.Value = RelationshipsNs;
                }
                else if (attribute.Name.NamespaceName == StrictRelationshipsNs)
                {
                    element.SetAttributeValue(Relationships + attribute.Name.LocalName, attribute.Value);
                    attribute.Remove();
                }
            }
        }
    }

    private XElement GetOrAdd(string name)
    {
        XElement? existing = Root.Element(Main + name);
        if (existing is not null)
            return existing;

        var element = new XElement(Main + name);
        Insert(element);
        return element;
    }

    private void Replace(string name, IEnumerable<XElement> replacements)
    {
        Root.Elements(Main + name).Remove();
        foreach (XElement replacement in replacements)
            Insert(replacement);
    }

    private void Insert(XElement element)
    {
        var index = Order(element);
        XElement? next = Root.Elements().FirstOrDefault(child => Order(child) > index);
        if (next is null)
            Root.Add(element);
        else
            next.AddBeforeSelf(element);

        static int Order(XElement child) => child.Name.Namespace == Main || child.Name == Compatibility + "AlternateContent"
            ? Array.IndexOf(ChildOrder, child.Name.LocalName)
            : -1;
    }

    private static void Flag(XElement element, XName name, bool value, bool? defaultValue = null)
        => element.SetAttributeValue(name, value == defaultValue ? null : value ? "1" : "0");

    internal void UpdateHyperlinks(XlsxWorksheet sheet, IEnumerable<string> oldIds, Action<string> remove, Func<Uri, string> add)
    {
        foreach (var id in oldIds)
            remove(id);

        if (!sheet.Hyperlinks.Any())
        {
            Replace("hyperlinks", []);
            return;
        }

        XElement hyperlinks = GetOrAdd("hyperlinks");
        hyperlinks.Elements(Main + "hyperlink").Remove();
        foreach (XlsxHyperlink link in sheet.Hyperlinks)
        {
            IXlsxCell cell = link.Cell ?? throw new InvalidOperationException("A worksheet hyperlink must be attached to a cell.");
            var element = new XElement(Main + "hyperlink", new XAttribute("ref", cell.Address.ToStringRelative()));
            if (link.IsExternal)
            {
                element.SetAttributeValue(Relationships + "id", add(link.ExternalAddress));
            }
            else
            {
                element.SetAttributeValue("location", link.InternalAddress);
                element.SetAttributeValue("display", cell.GetFormattedString());
            }
            if (!string.IsNullOrWhiteSpace(link.Tooltip))
                element.SetAttributeValue("tooltip", link.Tooltip);
            hyperlinks.Add(element);
        }
    }

    internal void UpdateTableReferences(IEnumerable<string> ids)
    {
        XElement tables = GetOrAdd("tableParts");
        tables.RemoveNodes();
        tables.Add(ids.Select(id => new XElement(Main + "tablePart", new XAttribute(Relationships + "id", id))));
        tables.SetAttributeValue("count", tables.Elements().Count());
    }

    internal void UpdateDrawingReferences(string? previousDrawingId, string? drawingId, string? legacyDrawingId)
    {
        XElement? drawing = previousDrawingId is null ? null : Root.Elements(Main + "drawing")
            .FirstOrDefault(element => (string?)element.Attribute(Relationships + "id") == previousDrawingId);
        if (drawingId is not null)
            (drawing ?? GetOrAdd("drawing")).SetAttributeValue(Relationships + "id", drawingId);
        else
            drawing?.Remove();

        Replace("legacyDrawing", string.IsNullOrEmpty(legacyDrawingId)
            ? []
            : [new XElement(Main + "legacyDrawing", new XAttribute(Relationships + "id", legacyDrawingId))]);
    }

    internal void Write(XmlWriter writer, Action writeData)
    {
        using XmlReader reader = Root.CreateReader();
        reader.MoveToContent();
        writer.WriteStartDocument(true);
        writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
        writer.WriteAttributes(reader, defattr: false);
        foreach (XNode node in Root.Nodes())
        {
            if (node is XElement data && data.Name == Main + "sheetData")
                writeData();
            else
                node.WriteTo(writer);
        }
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    internal void UpdateSettings(XlsxWorksheet sheet)
    {
        XElement properties = GetOrAdd("sheetPr");
        properties.Element(Main + "tabColor")?.Remove();
        if (sheet.TabColor.HasValue)
            properties.AddFirst(SpreadsheetColorXml.Write(Main + "tabColor", sheet.TabColor));

        XElement? outline = properties.Element(Main + "outlinePr");
        if (outline is null)
        {
            outline = new XElement(Main + "outlinePr");
            if (properties.Element(Main + "pageSetUpPr") is { } pageSetup)
                pageSetup.AddBeforeSelf(outline);
            else
                properties.Add(outline);
        }
        Flag(outline, "summaryBelow", sheet.Outline.SummaryVLocation == XlsxOutlineSummaryVLocation.Bottom);
        Flag(outline, "summaryRight", sheet.Outline.SummaryHLocation == XlsxOutlineSummaryHLocation.Right);
        if (properties.Element(Main + "pageSetUpPr") is null && (sheet.PageSetup.PagesTall > 0 || sheet.PageSetup.PagesWide > 0))
            properties.Add(new XElement(Main + "pageSetUpPr", new XAttribute("fitToPage", "1")));

        XElement dimension = GetOrAdd("dimension");
        if (dimension.Attribute("ref") is null)
        {
            XlsxCellsCollection cells = sheet.Internals.CellsCollection;
            dimension.SetAttributeValue("ref", cells.IsEmpty ? "A1" : "A1:" + XlsxHelper.GetColumnLetterFromNumber(cells.MaxColumnUsed) + cells.MaxRowUsed.ToInvariantString());
        }

        UpdateView(GetOrAdd("sheetViews"), sheet);
        XElement format = GetOrAdd("sheetFormatPr");
        format.SetAttributeValue("defaultRowHeight", sheet.RowHeight.SaveRound());
        Flag(format, "customHeight", sheet.RowHeightChanged, false);
        format.SetAttributeValue("defaultColWidth", sheet.ColumnWidthChanged ? ColumnWidth(sheet.ColumnWidth) : (double?)null);
        var columns = sheet.ColumnCount() > 0 ? sheet.GetMaxColumnOutline() : 0;
        var rows = sheet.RowCount() > 0 ? sheet.GetMaxRowOutline() : 0;
        format.SetAttributeValue("outlineLevelCol", columns > 0 ? columns : (int?)null);
        format.SetAttributeValue("outlineLevelRow", rows > 0 ? rows : (int?)null);

        GetOrAdd("sheetData");
        UpdateProtection(sheet);
        Replace("autoFilter", sheet.AutoFilter.IsEnabled ? [AutoFilterXml.Write(sheet.AutoFilter)] : []);

        if (sheet.Internals.MergedRanges.Count > 0)
        {
            XElement merges = GetOrAdd("mergeCells");
            merges.Elements(Main + "mergeCell").Remove();
            foreach (XlsxRange range in sheet.Internals.MergedRanges)
                merges.Add(new XElement(Main + "mergeCell", new XAttribute("ref", $"{range.RangeAddress.FirstAddress}:{range.RangeAddress.LastAddress}")));
            merges.SetAttributeValue("count", merges.Elements().Count());
        }
        else
        {
            Replace("mergeCells", []);
        }
    }

    private static void UpdateView(XElement sheetViews, XlsxWorksheet sheet)
    {
        XElement? view = sheetViews.Element(Main + "sheetView");
        if (view is null)
        {
            view = new XElement(Main + "sheetView", new XAttribute("workbookViewId", 0));
            sheetViews.AddFirst(view);
        }

        Flag(view, "tabSelected", sheet.TabSelected, false);
        Flag(view, "rightToLeft", sheet.RightToLeft, false);
        Flag(view, "showFormulas", sheet.ShowFormulas, false);
        Flag(view, "showGridLines", sheet.ShowGridLines, true);
        Flag(view, "showOutlineSymbols", sheet.ShowOutlineSymbols, true);
        Flag(view, "showRowColHeaders", sheet.ShowRowColHeaders, true);
        Flag(view, "showRuler", sheet.ShowRuler, true);
        Flag(view, "showWhiteSpace", sheet.ShowWhiteSpace, true);
        Flag(view, "showZeros", sheet.ShowZeros, true);

        XlsxSheetView settings = sheet.SheetView;
        view.SetAttributeValue("view", settings.View switch
        {
            XlsxSheetViewOptions.Normal => null,
            XlsxSheetViewOptions.PageBreakPreview => "pageBreakPreview",
            XlsxSheetViewOptions.PageLayout => "pageLayout",
            _ => throw new ArgumentOutOfRangeException(nameof(sheet)),
        });

        XElement? pane = view.Element(Main + "pane");
        var columns = settings.SplitColumn;
        var rows = settings.SplitRow;
        string? split = null;
        if (columns == 0 && rows == 0)
        {
            pane?.Remove();
            pane = null;
        }
        else
        {
            if (pane is null)
            {
                pane = new XElement(Main + "pane");
                view.AddFirst(pane);
            }
            split = columns == 0 ? "bottomLeft" : rows == 0 ? "topRight" : "bottomRight";
            pane.SetAttributeValue("state", "frozenSplit");
            pane.SetAttributeValue("xSplit", columns);
            pane.SetAttributeValue("ySplit", rows);
            pane.SetAttributeValue("activePane", split);
            pane.SetAttributeValue("topLeftCell", XlsxHelper.GetColumnLetterFromNumber(columns + 1) + (rows + 1).ToInvariantString());
        }

        XlsxAddress topLeft = settings.TopLeftCellAddress;
        view.SetAttributeValue("topLeftCell", !topLeft.IsValid || topLeft == new XlsxAddress(1, 1, false, false) ? null : topLeft.ToString());
        IXlsxRange? first = sheet.SelectedRanges.FirstOrDefault();
        if (first is not null || sheet.ActiveCell is not null)
        {
            view.Elements(Main + "selection").Remove();
            var activeCell = sheet.ActiveCell?.ToString() ?? first!.RangeAddress.FirstAddress.ToStringRelative(false);
            IEnumerable<string> references = sheet.SelectedRanges.Select(range => range.RangeAddress.FirstAddress.Equals(range.RangeAddress.LastAddress)
                ? range.RangeAddress.FirstAddress.ToStringRelative(false)
                : range.RangeAddress.ToStringRelative(false));
            var sequence = string.Join(" ", references.Prepend(activeCell).Distinct());
            var selection = new XElement(Main + "selection", new XAttribute("activeCell", activeCell), new XAttribute("sqref", sequence));
            if (pane is null)
            {
                view.AddFirst(selection);
            }
            else
            {
                var paneSelection = new XElement(selection);
                paneSelection.SetAttributeValue("pane", split);
                pane.AddAfterSelf(selection, paneSelection);
            }
        }

        view.SetAttributeValue("zoomScale", Zoom(settings.ZoomScale));
        view.SetAttributeValue("zoomScaleNormal", Zoom(settings.ZoomScaleNormal));
        view.SetAttributeValue("zoomScalePageLayoutView", Zoom(settings.ZoomScalePageLayoutView));
        view.SetAttributeValue("zoomScaleSheetLayoutView", Zoom(settings.ZoomScaleSheetLayoutView));

        static int? Zoom(int value) => value == 100 ? null : Math.Clamp(value, 10, 400);
    }

    private void UpdateProtection(XlsxWorksheet sheet)
    {
        XlsxSheetProtection protection = sheet.Protection;
        if (!protection.IsProtected)
        {
            Replace("sheetProtection", []);
            return;
        }

        XElement element = GetOrAdd("sheetProtection");
        Flag(element, "sheet", true);
        var simpleHash = protection.Algorithm == XlsxProtectionAlgorithm.Algorithm.SimpleHash;
        element.SetAttributeValue("password", simpleHash && !string.IsNullOrWhiteSpace(protection.PasswordHash) ? protection.PasswordHash : null);
        element.SetAttributeValue("algorithmName", simpleHash ? null : XlsxProtectionAlgorithm.ToDescription(protection.Algorithm));
        element.SetAttributeValue("hashValue", simpleHash ? null : protection.PasswordHash);
        element.SetAttributeValue("spinCount", simpleHash ? null : (uint?)protection.SpinCount);
        element.SetAttributeValue("saltValue", simpleHash ? null : protection.Base64EncodedSalt);

        SetDenied("formatCells", XlsxSheetProtectionElements.FormatCells, true);
        SetDenied("formatColumns", XlsxSheetProtectionElements.FormatColumns, true);
        SetDenied("formatRows", XlsxSheetProtectionElements.FormatRows, true);
        SetDenied("insertColumns", XlsxSheetProtectionElements.InsertColumns, true);
        SetDenied("insertRows", XlsxSheetProtectionElements.InsertRows, true);
        SetDenied("insertHyperlinks", XlsxSheetProtectionElements.InsertHyperlinks, true);
        SetDenied("deleteColumns", XlsxSheetProtectionElements.DeleteColumns, true);
        SetDenied("deleteRows", XlsxSheetProtectionElements.DeleteRows, true);
        SetDenied("sort", XlsxSheetProtectionElements.Sort, true);
        SetDenied("autoFilter", XlsxSheetProtectionElements.AutoFilter, true);
        SetDenied("pivotTables", XlsxSheetProtectionElements.PivotTables, true);
        SetDenied("scenarios", XlsxSheetProtectionElements.EditScenarios, true);
        SetDenied("objects", XlsxSheetProtectionElements.EditObjects, false);
        SetDenied("selectLockedCells", XlsxSheetProtectionElements.SelectLockedCells, false);
        SetDenied("selectUnlockedCells", XlsxSheetProtectionElements.SelectUnlockedCells, false);

        void SetDenied(string name, XlsxSheetProtectionElements flag, bool defaultValue)
            => Flag(element, name, !protection.AllowedElements.HasFlag(flag), defaultValue);
    }
}
