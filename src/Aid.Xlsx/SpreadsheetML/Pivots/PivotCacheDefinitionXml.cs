using System.Diagnostics;
using System.Xml;
using System.Xml.Linq;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Opc.Xml;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.SpreadsheetML.Pivots;

internal static class PivotCacheDefinitionXml
{
    internal static XElement Read(Stream stream)
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        var root = XElement.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        if (root.Name.LocalName != "pivotCacheDefinition" || root.Name.NamespaceName is not (OpenXmlConst.Main2006SsNs or OpenXmlConst.StrictSpreadsheetNs))
            throw PartStructureException.IncorrectElementFormat("pivotCacheDefinition");

        return root;
    }

    internal static XlsxPivotCache Load(XElement definition, string workbookRelId, XlsxWorkbook workbook)
    {
        var pivotCache = workbook.PivotCachesInternal.Add(ReadSource(definition));
        if (string.IsNullOrWhiteSpace(pivotCache.WorkbookCacheRelId))
            pivotCache.WorkbookCacheRelId = workbookRelId;

        if ((uint?)definition.Attribute("missingItemsLimit") is { } limit)
        {
            pivotCache.ItemsToRetainPerField = limit switch
            {
                0 => XlsxItemsToRetain.None,
                XlsxHelper.MaxRowNumber => XlsxItemsToRetain.Max,
                _ => XlsxItemsToRetain.Automatic
            };
        }

        var ns = definition.Name.Namespace;
        if (definition.Element(ns + "cacheFields") is { } fields)
        {
            foreach (var field in fields.Elements(ns + "cacheField"))
            {
                var name = Required(field, "name");
                if (!pivotCache.ContainsField(name))
                {
                    var items = field.Element(ns + "sharedItems");
                    pivotCache.AddCachedField(name, new XlsxPivotCacheValues(ReadSharedItems(items), ReadStats(items)));
                }
            }
        }

        pivotCache.SaveSourceData = (bool?)definition.Attribute("saveData") ?? true;
        return pivotCache;
    }

    internal static IXlsxPivotSource ReadSource(XElement definition)
    {
        var ns = definition.Name.Namespace;
        var source = definition.Element(ns + "cacheSource") ?? throw PartStructureException.RequiredElementIsMissing("cacheSource");
        var type = Required(source, "type");
        return type switch
        {
            "worksheet" => ReadWorksheetSource(source.Element(ns + "worksheetSource") ?? throw PartStructureException.RequiredElementIsMissing("worksheetSource")),
            "external" => new XlsxPivotSourceConnection((uint?)source.Attribute("connectionId") ?? throw PartStructureException.MissingAttribute("connectionId")),
            "consolidation" => ReadConsolidation(source.Element(ns + "consolidation") ?? throw PartStructureException.RequiredElementIsMissing("consolidation")),
            "scenario" => new XlsxPivotSourceScenario(),
            _ => throw PartStructureException.InvalidAttributeValue(type)
        };
    }

    private static IXlsxPivotSource ReadWorksheetSource(XElement element)
    {
        var relId = (string?)element.Attribute(RelationshipId(element.Name.Namespace));
        if ((string?)element.Attribute("name") is { } name)
            return relId is null ? new XlsxPivotSourceReference(name) : new XlsxPivotSourceExternalWorkbook(relId, name);

        if (ReadArea(element) is { } area)
            return relId is null ? new XlsxPivotSourceReference(area) : new XlsxPivotSourceExternalWorkbook(relId, area);

        throw PartStructureException.IncorrectElementFormat("worksheetSource");
    }

    private static XlsxPivotSourceConsolidation ReadConsolidation(XElement consolidation)
    {
        var ns = consolidation.Name.Namespace;
        var pages = consolidation.Element(ns + "pages")?.Elements(ns + "page")
            .Select(page => new XlsxPivotCacheSourceConsolidationPage(page.Elements(ns + "pageItem").Select(item => Required(item, "name")).ToList()))
            .ToList() ?? [];
        var ranges = consolidation.Element(ns + "rangeSets") ?? throw PartStructureException.RequiredElementIsMissing("rangeSets");
        var rangeSets = new List<XlsxPivotCacheSourceConsolidationRangeSet>();
        foreach (var range in ranges.Elements(ns + "rangeSet"))
        {
            uint?[] indexes = [(uint?)range.Attribute("i1"), (uint?)range.Attribute("i2"), (uint?)range.Attribute("i3"), (uint?)range.Attribute("i4")];
            for (var i = 0; i < indexes.Length; ++i)
            {
                if (indexes[i] is { } index && (i >= pages.Count || index >= pages[i].PageItems.Count))
                    throw PartStructureException.InvalidAttributeValue();
            }

            var name = (string?)range.Attribute("name");
            var area = name is null ? ReadArea(range) ?? throw PartStructureException.IncorrectElementFormat("rangeSet") : (SheetArea?)null;
            rangeSets.Add(new XlsxPivotCacheSourceConsolidationRangeSet
            {
                Indexes = indexes,
                RelId = (string?)range.Attribute(RelationshipId(ns)),
                TableOrName = name,
                Area = area
            });
        }

        if (rangeSets.Count == 0)
            throw PartStructureException.IncorrectElementsCount();

        return new XlsxPivotSourceConsolidation
        {
            AutoPage = (bool?)consolidation.Attribute("autoPage") ?? true,
            Pages = pages,
            RangeSets = rangeSets
        };
    }

    private static SheetArea? ReadArea(XElement element)
    {
        return (string?)element.Attribute("sheet") is { } sheet &&
               (string?)element.Attribute("ref") is { } reference && Area.TryParse(reference, out var area)
            ? new SheetArea(sheet, area)
            : null;
    }

    private static XlsxPivotCacheValuesStats ReadStats(XElement? items) => new(
        (bool?)items?.Attribute("containsBlank") ?? false,
        (bool?)items?.Attribute("containsNumber") ?? false,
        (bool?)items?.Attribute("containsInteger") ?? false,
        (double?)items?.Attribute("minValue"),
        (double?)items?.Attribute("maxValue"),
        (bool?)items?.Attribute("containsString") ?? true,
        (bool?)items?.Attribute("longText") ?? false,
        (bool?)items?.Attribute("containsDate") ?? false,
        (DateTime?)items?.Attribute("minDate"),
        (DateTime?)items?.Attribute("maxDate"));

    private static XlsxPivotCacheSharedItems ReadSharedItems(XElement? element)
    {
        var items = new XlsxPivotCacheSharedItems();
        if (element is null)
            return items;

        var ns = element.Name.Namespace;
        foreach (var item in element.Elements().Where(item => item.Name.Namespace == ns))
        {
            switch (item.Name.LocalName)
            {
                case "m":
                    items.AddMissing();
                    break;
                case "n":
                    items.AddNumber((double?)item.Attribute("v") ?? throw PartStructureException.MissingAttribute("v"));
                    break;
                case "b":
                    items.AddBoolean((bool?)item.Attribute("v") ?? throw PartStructureException.MissingAttribute("v"));
                    break;
                case "e":
                    if (!XlsxErrorParser.TryParseError(Required(item, "v"), out var error))
                        throw PartStructureException.InvalidAttributeFormat();
                    items.AddError(error);
                    break;
                case "s":
                    items.AddString(Required(item, "v"));
                    break;
                case "d":
                    items.AddDateTime((DateTime?)item.Attribute("v") ?? throw PartStructureException.MissingAttribute("v"));
                    break;
                default:
                    throw PartStructureException.ExpectedElementNotFound();
            }
        }

        return items;
    }

    internal static void Write(Stream stream, XElement? definition, XlsxPivotCache cache, string recordsRelId)
    {
        XNamespace ns = definition?.Name.Namespace ?? OpenXmlConst.Main2006SsNs;
        var rel = RelationshipId(ns);
        definition ??= new XElement(ns + "pivotCacheDefinition",
            new XAttribute(XNamespace.Xmlns + "r", rel.NamespaceName), new XAttribute(XNamespace.Xmlns + "x", ns.NamespaceName),
            new XAttribute("refreshOnLoad", "1"));

        if (cache.SaveSourceData || definition.Attribute(rel) is not null)
            definition.SetAttributeValue(rel, recordsRelId);
        SetVersion("createdVersion", XlsxConstants.PivotTable.CreatedVersion);
        SetVersion("refreshedVersion", XlsxConstants.PivotTable.RefreshedVersion);
        SetVersion("minRefreshableVersion", 3);
        definition.SetAttributeValue("saveData", cache.SaveSourceData ? "1" : "0");
        definition.SetAttributeValue("missingItemsLimit", cache.ItemsToRetainPerField switch
        {
            XlsxItemsToRetain.None => 0,
            XlsxItemsToRetain.Max => XlsxHelper.MaxRowNumber,
            _ => (uint?)definition.Attribute("missingItemsLimit")
        });

        var source = WriteSource(ns, cache.Source);
        if (definition.Element(ns + "cacheSource") is { } existingSource)
            existingSource.ReplaceWith(source);
        else
            definition.AddFirst(source);

        var fields = definition.Element(ns + "cacheFields");
        if (fields is null)
        {
            fields = new XElement(ns + "cacheFields");
            source.AddAfterSelf(fields);
        }

        fields.RemoveNodes();
        if (fields.Attribute("count") is not null)
            fields.SetAttributeValue("count", cache.FieldCount);
        for (var index = 0; index < cache.FieldCount; ++index)
        {
            var values = cache.GetFieldValues(index);
            var items = WriteSharedItems(ns, values, cache.GetFieldSharedItems(index));
            fields.Add(new XElement(ns + "cacheField", new XAttribute("name", cache.FieldNames[index]), items));
        }

        new XDocument(definition).Save(stream, System.Xml.Linq.SaveOptions.DisableFormatting);
        return;

        void SetVersion(string name, byte minimum) => definition.SetAttributeValue(name, Math.Max(minimum, (int?)definition.Attribute(name) ?? 0));
    }

    private static XElement WriteSource(XNamespace ns, IXlsxPivotSource source)
    {
        var cacheSource = new XElement(ns + "cacheSource");
        switch (source)
        {
            case XlsxPivotSourceReference local:
                cacheSource.SetAttributeValue("type", "worksheet");
                cacheSource.Add(WriteReference(ns + "worksheetSource", local.UsesName ? local.Name : null, local.Area, null));
                break;
            case XlsxPivotSourceExternalWorkbook external:
                cacheSource.SetAttributeValue("type", "worksheet");
                cacheSource.Add(WriteReference(ns + "worksheetSource", external.UsesName ? external.TableOrName : null, external.Area, external.RelId));
                break;
            case XlsxPivotSourceConnection connection:
                cacheSource.SetAttributeValue("type", "external");
                cacheSource.SetAttributeValue("connectionId", connection.ConnectionId);
                break;
            case XlsxPivotSourceConsolidation consolidation:
                cacheSource.SetAttributeValue("type", "consolidation");
                var pages = consolidation.Pages.Count > 0
                    ? new XElement(ns + "pages", consolidation.Pages.Select(page => new XElement(ns + "page",
                        page.PageItems.Select(item => new XElement(ns + "pageItem", new XAttribute("name", item))))))
                    : null;
                var ranges = new XElement(ns + "rangeSets");
                foreach (var range in consolidation.RangeSets)
                {
                    var element = WriteReference(ns + "rangeSet", range.UsesName ? range.TableOrName : null, range.Area, range.RelId);
                    for (var index = 0; index < Math.Min(range.Indexes.Count, 4); ++index)
                        element.SetAttributeValue($"i{index + 1}", range.Indexes[index]);
                    ranges.Add(element);
                }

                cacheSource.Add(new XElement(ns + "consolidation", Flag("autoPage", consolidation.AutoPage), pages, ranges));
                break;
            case XlsxPivotSourceScenario:
                cacheSource.SetAttributeValue("type", "scenario");
                break;
            default:
                throw new UnreachableException();
        }

        return cacheSource;
    }

    private static XElement WriteReference(XName element, string? name, SheetArea? area, string? relId)
    {
        var reference = new XElement(element);
        reference.SetAttributeValue(RelationshipId(element.Namespace), relId);
        if (name is not null)
            reference.SetAttributeValue("name", name);
        else if (area is { } sheetArea)
        {
            reference.SetAttributeValue("ref", sheetArea.Area.ToString());
            reference.SetAttributeValue("sheet", sheetArea.Name);
        }

        return reference;
    }

    private static XElement WriteSharedItems(XNamespace ns, XlsxPivotCacheValues values, XlsxPivotCacheSharedItems shared)
    {
        var stats = values.Stats;
        var typesCount = (stats.ContainsNumber ? 1 : 0) + (stats.ContainsString ? 1 : 0) + (stats.ContainsDate ? 1 : 0);
        var items = new XElement(ns + "sharedItems",
            Flag("containsBlank", stats.ContainsBlank, false),
            Flag("containsDate", stats.ContainsDate, false),
            Flag("containsMixedTypes", typesCount > 1, false),
            Flag("containsNonDate", stats.ContainsString || stats.ContainsNumber, true),
            Flag("containsSemiMixedTypes", stats.ContainsString || stats.ContainsBlank, true),
            Flag("containsString", stats.ContainsString, true),
            Flag("longText", stats.LongText, false));

        // A field with dates interprets numeric values as OLE Automation dates.
        if (stats.ContainsDate)
        {
            var minDate = DateBound(stats.MinDate, stats.MinValue, takeMinimum: true);
            var maxDate = DateBound(stats.MaxDate, stats.MaxValue, takeMinimum: false);
            items.SetAttributeValue("minDate", minDate is { } min ? DateText(min) : null);
            items.SetAttributeValue("maxDate", maxDate is { } max ? DateText(max) : null);
        }
        else if (stats.ContainsNumber)
        {
            items.SetAttributeValue("containsNumber", "1");
            items.Add(Flag("containsInteger", stats.ContainsInteger, false));
            items.SetAttributeValue("minValue", stats.MinValue);
            items.SetAttributeValue("maxValue", stats.MaxValue);
        }

        if (values.SharedCount != 0)
            items.SetAttributeValue("count", values.SharedCount);

        foreach (var value in shared.GetCellValues())
        {
            items.Add(value.Type switch
            {
                XlsxDataType.Blank => new XElement(ns + "m"),
                XlsxDataType.Boolean => new XElement(ns + "b", Flag("v", value.GetBoolean())),
                XlsxDataType.Number => Item("n", value.GetNumber()),
                XlsxDataType.Text => Item("s", value.GetText()),
                XlsxDataType.Error => Item("e", value.GetError().ToDisplayString()),
                XlsxDataType.DateTime => Item("d", DateText(value.GetDateTime())),
                XlsxDataType.TimeSpan => Item("d", DateText(DateTime.FromOADate(value.GetUnifiedNumber()))),
                _ => throw new InvalidOperationException()
            });
        }

        return items;

        XElement Item(string name, object value) => new(ns + name, new XAttribute("v", value));
    }

    private static DateTime? DateBound(DateTime? date, double? serial, bool takeMinimum)
    {
        if (serial is not { } number)
            return date;

        var serialDate = DateTime.FromOADate(number);
        if (date is not { } existing)
            return serialDate;

        return takeMinimum == (existing < serialDate) ? existing : serialDate;
    }

    private static string DateText(DateTime value) => XmlConvert.ToString(value, XmlDateTimeSerializationMode.RoundtripKind);

    private static XAttribute? Flag(string name, bool value, bool? defaultValue = null) =>
        value == defaultValue ? null : new XAttribute(name, value ? "1" : "0");

    private static XName RelationshipId(XNamespace ns) => XName.Get("id",
        ns == OpenXmlConst.StrictSpreadsheetNs ? OpenXmlConst.StrictRelationshipsNs : OpenXmlConst.RelationshipsNs);

    private static string Required(XElement element, string name) =>
        (string?)element.Attribute(name) ?? throw PartStructureException.MissingAttribute(name);
}
