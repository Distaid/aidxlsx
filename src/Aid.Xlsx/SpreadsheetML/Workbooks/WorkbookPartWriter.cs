using System.Xml.Linq;
using Aid.Xlsx.DefinedNames;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.PivotTables;
using Aid.Xlsx.Protection;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.SpreadsheetML;

internal static class WorkbookPartWriter
{
    private static readonly string[] ChildOrder =
    [
        "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
        "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews",
        "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"
    ];

    internal static XElement CreateRoot()
    {
        XNamespace ns = OpenXmlConst.Main2006SsNs;
        return new XElement(ns + "workbook",
            new XAttribute(XNamespace.Xmlns + "r", OpenXmlConst.RelationshipsNs),
            new XAttribute(XNamespace.Xmlns + "x", ns));
    }

    internal static void GenerateContent(XElement root, XlsxWorkbook workbook, Workbooks.SaveOptions options, XlsxWorkbook.SaveContext context)
    {
        var ns = root.Name.Namespace;
        var relationship = RelationshipId(root);
        if (root.GetPrefixOfNamespace(relationship.Namespace) is null)
            root.SetAttributeValue(XNamespace.Xmlns + "r", relationship.NamespaceName);

        var properties = GetOrAdd(root, "workbookPr");
        properties.SetAttributeValue("codeName", (string?)properties.Attribute("codeName") ?? "ThisWorkbook");
        properties.SetAttributeValue("date1904", workbook.Use1904DateSystem ? "1" : null);
        if (options.FilterPrivacy is { } privacy)
            properties.SetAttributeValue("filterPrivacy", privacy ? "1" : null);

        var sharing = GetOrAdd(root, "fileSharing");
        sharing.SetAttributeValue("readOnlyRecommended", workbook.FileSharing.ReadOnlyRecommended ? "1" : null);
        sharing.SetAttributeValue("userName", string.IsNullOrWhiteSpace(workbook.FileSharing.UserName) ? null : workbook.FileSharing.UserName);
        if (!sharing.HasElements && !sharing.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration))
            sharing.Remove();

        WriteProtection(root, workbook);

        var sheets = GetOrAdd(root, "sheets");
        var sheetById = sheets.Elements(ns + "sheet").ToDictionary(sheet => WorkbookXml.ReadSheet(sheet).SheetId);
        var ordered = new SortedDictionary<int, XElement>();
        foreach (var unsupported in workbook.UnsupportedSheets)
            ordered.Add(unsupported.Position, sheetById[unsupported.SheetId]);

        foreach (var worksheet in workbook.WorksheetsInternal.Cast<XlsxWorksheet>().OrderBy(sheet => sheet.Position))
        {
            if (!sheetById.TryGetValue(worksheet.SheetId, out var sheet))
                sheet = new XElement(ns + "sheet", new XAttribute("name", worksheet.Name), new XAttribute("sheetId", worksheet.SheetId));

            var originalId = WorkbookXml.ReadSheet(sheet).Id;
            var id = string.IsNullOrWhiteSpace(originalId) ? worksheet.RelId : originalId;
            if (string.IsNullOrWhiteSpace(id))
                id = context.RelIdGenerator.GetNext(XlsxWorkbook.RelType.Workbook);
            worksheet.RelId = id;

            sheet.SetAttributeValue("name", worksheet.Name);
            sheet.SetAttributeValue("state", worksheet.Visibility switch
            {
                XlsxWorksheetVisibility.Visible => null,
                XlsxWorksheetVisibility.Hidden => "hidden",
                XlsxWorksheetVisibility.VeryHidden => "veryHidden",
                _ => throw new ArgumentOutOfRangeException(nameof(workbook), "Invalid worksheet visibility.")
            });
            sheet.SetAttributeValue(relationship, id);
            ordered.Add(worksheet.Position, sheet);
        }

        sheets.Elements(ns + "sheet").Remove();
        sheets.Add(ordered.Values);
        WriteView(root, workbook, ordered.Values);
        WriteDefinedNames(root, workbook, ordered.Values);
        WriteCalculation(root, workbook);
    }

    internal static void WritePivotCaches(XElement root, IReadOnlyList<XlsxPivotCache> caches)
    {
        var ns = root.Name.Namespace;
        root.Element(ns + "pivotCaches")?.Remove();
        if (caches.Count == 0)
            return;

        var list = GetOrAdd(root, "pivotCaches");
        var relationship = RelationshipId(root);
        for (var index = 0; index < caches.Count; index++)
        {
            var cache = caches[index];
            cache.CacheId = (uint)index;
            var id = cache.WorkbookCacheRelId ?? throw new InvalidOperationException("A pivot cache has no workbook relationship.");
            list.Add(new XElement(ns + "pivotCache", new XAttribute("cacheId", cache.CacheId), new XAttribute(relationship, id)));
        }
    }

    private static void WriteProtection(XElement root, XlsxWorkbook workbook)
    {
        var protection = workbook.Protection;
        if (!protection.IsProtected)
        {
            root.Element(root.Name.Namespace + "workbookProtection")?.Remove();
            return;
        }

        var element = GetOrAdd(root, "workbookProtection");
        var simpleHash = protection.Algorithm == XlsxProtectionAlgorithm.Algorithm.SimpleHash;
        element.SetAttributeValue("workbookPassword", simpleHash && !string.IsNullOrWhiteSpace(protection.PasswordHash) ? protection.PasswordHash : null);
        element.SetAttributeValue("lockStructure", protection.AllowedElements.HasFlag(XlsxWorkbookProtectionElements.Structure) ? null : "1");
        element.SetAttributeValue("lockWindows", protection.AllowedElements.HasFlag(XlsxWorkbookProtectionElements.Windows) ? null : "1");
        element.SetAttributeValue("workbookAlgorithmName", simpleHash ? null : XlsxProtectionAlgorithm.ToDescription(protection.Algorithm));
        element.SetAttributeValue("workbookHashValue", simpleHash ? null : protection.PasswordHash);
        element.SetAttributeValue("workbookSaltValue", simpleHash ? null : protection.Base64EncodedSalt);
        element.SetAttributeValue("workbookSpinCount", simpleHash ? null : protection.SpinCount);
    }

    private static void WriteView(XElement root, XlsxWorkbook workbook, IEnumerable<XElement> sheets)
    {
        var unsupported = workbook.UnsupportedSheets.Select(sheet => sheet.SheetId).ToHashSet();
        var firstVisible = sheets.Select(WorkbookXml.ReadSheet).Where(sheet => !unsupported.Contains(sheet.SheetId))
            .TakeWhile(sheet => sheet.Visibility != XlsxWorksheetVisibility.Visible).Count();
        var active = workbook.UnsupportedSheets.Where(sheet => sheet.IsActive).Select(sheet => (int?)sheet.Position - 1).FirstOrDefault();
        if (!active.HasValue)
        {
            var worksheets = workbook.WorksheetsInternal.Cast<XlsxWorksheet>();
            active = (worksheets.FirstOrDefault(sheet => sheet.TabActive)?.Position - 1)
                ?? (worksheets.LastOrDefault(sheet => sheet.TabSelected)?.Position - 1)
                ?? firstVisible;
        }

        var bookViews = GetOrAdd(root, "bookViews");
        var name = root.Name.Namespace + "workbookView";
        var view = bookViews.Element(name);
        if (view is null)
        {
            view = new XElement(name);
            bookViews.Add(view);
        }

        view.SetAttributeValue("firstSheet", firstVisible);
        view.SetAttributeValue("activeTab", active);
    }

    private static void WriteDefinedNames(XElement root, XlsxWorkbook workbook, IEnumerable<XElement> sheets)
    {
        var ns = root.Name.Namespace;
        var indexes = sheets.Select((sheet, index) => (Sheet: WorkbookXml.ReadSheet(sheet), Index: index))
            .ToDictionary(item => item.Sheet.SheetId, item => item.Index);
        var names = GetOrAdd(root, "definedNames");
        names.Elements(ns + "definedName").Remove();
        foreach (var worksheet in workbook.WorksheetsInternal.Cast<XlsxWorksheet>())
        {
            var sheetIndex = indexes[worksheet.SheetId];
            var prefix = worksheet.Name.EscapeSheetName() + "!";
            var setup = worksheet.PageSetup;
            if (setup.PrintAreas.Any())
                AddName("_xlnm.Print_Area", string.Join(",", setup.PrintAreas.Select(range => prefix + FixedRange(range.RangeAddress))), sheetIndex);

            if (worksheet.AutoFilter.IsEnabled)
                AddName("_xlnm._FilterDatabase", prefix + FixedRange(worksheet.AutoFilter.Range.RangeAddress), sheetIndex, hidden: true);

            foreach (var name in worksheet.DefinedNames.Cast<XlsxDefinedName>().Where(name => name.Name != "_xlnm._FilterDatabase"))
                AddName(name.Name, name.ToString(), sheetIndex, !name.Visible, name.Comment);

            var titles = new List<string>(2);
            if (setup.FirstColumnToRepeatAtLeft > 0)
                titles.Add(prefix + XlsxHelper.GetColumnLetterFromNumber(setup.FirstColumnToRepeatAtLeft) + ":" + XlsxHelper.GetColumnLetterFromNumber(setup.LastColumnToRepeatAtLeft));
            if (setup.FirstRowToRepeatAtTop > 0)
                titles.Add(prefix + FormattableString.Invariant($"{setup.FirstRowToRepeatAtTop}:{setup.LastRowToRepeatAtTop}"));
            if (titles.Count > 0)
                AddName("_xlnm.Print_Titles", string.Join(",", titles), sheetIndex);
        }

        foreach (var name in workbook.DefinedNamesInternal)
            AddName(name.Name, name.RefersTo, hidden: !name.Visible, comment: name.Comment);

        void AddName(string name, string text, int? sheetIndex = null, bool hidden = false, string? comment = null)
        {
            names.Add(new XElement(ns + "definedName", new XAttribute("name", name),
                string.IsNullOrWhiteSpace(comment) ? null : new XAttribute("comment", comment),
                sheetIndex.HasValue ? new XAttribute("localSheetId", sheetIndex.Value) : null,
                hidden ? new XAttribute("hidden", "1") : null, text));
        }

        static string FixedRange(IXlsxRangeAddress range) => range.FirstAddress.ToStringFixed(XlsxReferenceStyle.A1) + ":" + range.LastAddress.ToStringFixed(XlsxReferenceStyle.A1);
    }

    private static void WriteCalculation(XElement root, XlsxWorkbook workbook)
    {
        var calculation = root.Element(root.Name.Namespace + "calcPr");
        if (calculation is null)
        {
            calculation = GetOrAdd(root, "calcPr");
            calculation.SetAttributeValue("calcId", 125725);
        }

        calculation.SetAttributeValue("calcMode", workbook.CalculateMode switch
        {
            XlsxCalculateMode.Default => null,
            XlsxCalculateMode.Auto => "auto",
            XlsxCalculateMode.AutoNoTable => "autoNoTable",
            XlsxCalculateMode.Manual => "manual",
            _ => throw new ArgumentOutOfRangeException(nameof(workbook), "Invalid workbook calculation mode.")
        });
        calculation.SetAttributeValue("refMode", workbook.ReferenceStyle switch
        {
            XlsxReferenceStyle.Default => null,
            XlsxReferenceStyle.A1 => "A1",
            XlsxReferenceStyle.R1C1 => "R1C1",
            _ => throw new ArgumentOutOfRangeException(nameof(workbook), "Invalid workbook reference style.")
        });
        if (workbook.FullCalculationOnLoad) calculation.SetAttributeValue("fullCalcOnLoad", "1");
        if (workbook.FullPrecision) calculation.SetAttributeValue("fullPrecision", "1");
        if (workbook.CalculationOnSave) calculation.SetAttributeValue("calcOnSave", "1");
        if (workbook.ForceFullCalculation) calculation.SetAttributeValue("forceFullCalc", "1");
    }

    private static XName RelationshipId(XElement root) => XName.Get("id",
        root.Name.NamespaceName == OpenXmlConst.StrictSpreadsheetNs ? OpenXmlConst.StrictRelationshipsNs : OpenXmlConst.RelationshipsNs);

    private static XElement GetOrAdd(XElement root, string localName)
    {
        var ns = root.Name.Namespace;
        var existing = root.Element(ns + localName);
        if (existing is not null)
            return existing;

        var position = Array.IndexOf(ChildOrder, localName);
        var next = root.Elements().FirstOrDefault(element => element.Name.Namespace == ns && Array.IndexOf(ChildOrder, element.Name.LocalName) > position);
        var added = new XElement(ns + localName);
        if (next is null)
            root.Add(added);
        else
            next.AddBeforeSelf(added);
        return added;
    }
}
