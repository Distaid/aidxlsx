using System.Globalization;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Opc.Packaging;

namespace Aid.Xlsx.SpreadsheetML.Workbooks;

internal sealed partial class WorkbookPackage(OpcPackage package)
{
    private const string TransitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
    private const string SpreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.";
    private readonly Dictionary<string, int> _pivotPartCounts = package.Parts.Values
        .Where(part => part.ContentType.StartsWith(SpreadsheetContentType + "pivot", StringComparison.Ordinal))
        .GroupBy(part => part.ContentType, StringComparer.Ordinal)
        .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    internal OpcPart? WorkbookPart => RelatedParts(package.Relationships, "officeDocument").FirstOrDefault().Part;

    internal OpcPart AddWorkbook(string contentType) =>
        AddPart(null, "rId1", "officeDocument", "/xl/workbook.xml", contentType, numbered: false);

    internal OpcPart AddPart(OpcPart source, string id, string type)
    {
        (var name, var contentType, var numbered) = type switch
        {
            "worksheet" => ("/xl/worksheets/sheet.xml", SpreadsheetContentType + "worksheet+xml", true),
            "styles" => ("/xl/styles.xml", SpreadsheetContentType + "styles+xml", false),
            "sharedStrings" => ("/xl/sharedStrings.xml", SpreadsheetContentType + "sharedStrings+xml", false),
            "theme" => ("/xl/theme/theme.xml", "application/vnd.openxmlformats-officedocument.theme+xml", true),
            "calcChain" => ("/xl/calcChain.xml", SpreadsheetContentType + "calcChain+xml", false),
            "table" => ("/xl/tables/table.xml", SpreadsheetContentType + "table+xml", true),
            "comments" => ("/xl/comments.xml", SpreadsheetContentType + "comments+xml", true),
            "drawing" => ("/xl/drawings/drawing.xml", "application/vnd.openxmlformats-officedocument.drawing+xml", true),
            "vmlDrawing" => ("/xl/drawings/vmlDrawing.vml", "application/vnd.openxmlformats-officedocument.vmlDrawing", true),
            "pivotTable" => ("/xl/pivotTables/pivotTable.xml", SpreadsheetContentType + "pivotTable+xml", false),
            "pivotCacheDefinition" => ("/pivotCache/pivotCacheDefinition.xml", SpreadsheetContentType + "pivotCacheDefinition+xml", true),
            "pivotCacheRecords" => (source.PartName[..(source.PartName.LastIndexOf('/') + 1)] + "pivotCacheRecords.xml",
                SpreadsheetContentType + "pivotCacheRecords+xml", true),
            _ => throw new ArgumentOutOfRangeException(nameof(type)),
        };
        return AddPart(source, id, type, name, contentType, numbered);
    }

    internal OpcPart AddImage(OpcPart source, string id, XlsxPictureFormat format)
    {
        (var extension, var contentType) = format switch
        {
            XlsxPictureFormat.Bmp => (".bmp", "image/bmp"),
            XlsxPictureFormat.Gif => (".gif", "image/gif"),
            XlsxPictureFormat.Png => (".png", "image/png"),
            XlsxPictureFormat.Tiff => (".tiff", "image/tiff"),
            XlsxPictureFormat.Icon => (".ico", "image/x-icon"),
            XlsxPictureFormat.Pcx => (".pcx", "image/x-pcx"),
            XlsxPictureFormat.Jpeg => (".jpeg", "image/jpeg"),
            XlsxPictureFormat.Emf => (".emf", "image/x-emf"),
            XlsxPictureFormat.Wmf => (".wmf", "image/x-wmf"),
            XlsxPictureFormat.Webp => (".webp", "image/webp"),
            _ => throw new ArgumentOutOfRangeException(nameof(format)),
        };
        return AddPart(source, id, "image", "/xl/media/image" + extension, contentType, numbered: true);
    }

    internal static void Link(OpcPart source, string id, string type, OpcPart target)
    {
        Uri relativeUri = new Uri("https://package.invalid" + source.PartName)
            .MakeRelativeUri(new Uri("https://package.invalid" + target.PartName));
        source.Relationships.Add(id, TransitionalRelationships + type, relativeUri.OriginalString);
    }

    internal static void LinkHyperlink(OpcPart source, string id, Uri target) =>
        source.Relationships.Add(id, TransitionalRelationships + "hyperlink", target.OriginalString, isExternal: true);

    internal void Unlink(OpcPart source, string id)
    {
        if (!source.Relationships.TryGet(id, out OpcRelationship? relationship))
            return;

        source.Relationships.Remove(id);
        if (Resolve(relationship) is not { } target)
            return;

        var candidates = new HashSet<OpcPart>();
        var pending = new Stack<OpcPart>();
        pending.Push(target);
        while (pending.TryPop(out OpcPart? candidate))
        {
            if (!candidates.Add(candidate))
                continue;

            foreach (OpcRelationship child in candidate.Relationships)
                if (IsManagedChild(child) && Resolve(child) is { } childPart)
                    pending.Push(childPart);
        }

        // A part linked by any surviving source remains shared, together with its descendants.
        foreach (OpcRelationship rootRelationship in package.Relationships)
            KeepTargets(rootRelationship);

        foreach (OpcPart part in package.Parts.Values)
            if (!candidates.Contains(part))
                foreach (OpcRelationship child in part.Relationships)
                    KeepTargets(child);

        foreach (OpcPart part in candidates)
            package.RemovePart(part.PartName);

        void KeepTargets(OpcRelationship link)
        {
            if (Resolve(link) is { } linkedPart)
                pending.Push(linkedPart);

            while (pending.TryPop(out OpcPart? candidate))
            {
                if (!candidates.Remove(candidate))
                    continue;

                foreach (OpcRelationship child in candidate.Relationships)
                    if (Resolve(child) is { } childPart)
                        pending.Push(childPart);
            }
        }
    }

    private static bool IsManagedChild(OpcRelationship relationship) =>
        IsOfficeRelationship(relationship, "table") || IsOfficeRelationship(relationship, "comments") ||
        IsOfficeRelationship(relationship, "vmlDrawing") || IsOfficeRelationship(relationship, "drawing") ||
        IsOfficeRelationship(relationship, "image") || IsOfficeRelationship(relationship, "pivotTable") ||
        IsOfficeRelationship(relationship, "pivotCacheDefinition") || IsOfficeRelationship(relationship, "pivotCacheRecords");

    private OpcPart AddPart(OpcPart? source, string id, string type, string name, string contentType, bool numbered)
    {
        var partName = name;
        var isPivot = contentType.StartsWith(SpreadsheetContentType + "pivot", StringComparison.Ordinal);
        var index = isPivot ? _pivotPartCounts.GetValueOrDefault(contentType) + 1 : 1;
        if (numbered || index > 1 || package.Parts.ContainsKey(partName))
        {
            var dot = name.LastIndexOf('.');
            for (; ; index++)
            {
                partName = name.Insert(dot, index.ToString(CultureInfo.InvariantCulture));
                if (!package.Parts.ContainsKey(partName))
                    break;
            }
        }

        if (isPivot)
            _pivotPartCounts[contentType] = index;
        OpcPart part = package.AddPart(partName, contentType);
        if (source is null)
            package.Relationships.Add(id, TransitionalRelationships + type, partName[1..]);
        else
            Link(source, id, type, part);
        return part;
    }

    internal static bool IsOfficeRelationship(OpcRelationship relationship, string type) =>
        HasType(relationship.Type, TransitionalRelationships, type) || HasType(relationship.Type, StrictRelationships, type);

    internal (string Id, OpcPart Part) FirstRelatedPart(OpcPart source, string type) =>
        RelatedParts(source, type).FirstOrDefault();

    internal IEnumerable<(string Id, OpcPart Part)> RelatedParts(OpcPart source, string type) =>
        RelatedParts(source.Relationships, type);

    internal OpcPart? RelatedPart(OpcPart source, string id, string type) =>
        source.Relationships.TryGet(id, out OpcRelationship? relationship) && IsOfficeRelationship(relationship, type)
            ? Resolve(relationship)
            : null;

    private IEnumerable<(string Id, OpcPart Part)> RelatedParts(OpcRelationshipCollection relationships, string type)
    {
        foreach (OpcRelationship relationship in relationships)
        {
            if (IsOfficeRelationship(relationship, type) && Resolve(relationship) is { } target)
                yield return (relationship.Id, target);
        }
    }

    private OpcPart? Resolve(OpcRelationship relationship) =>
        relationship.TargetPartName is { } name && package.TryGetPart(name, out OpcPart? part) ? part : null;

    private static bool HasType(string qualifiedType, string prefix, string type) =>
        qualifiedType.StartsWith(prefix, StringComparison.Ordinal) && qualifiedType.AsSpan(prefix.Length).SequenceEqual(type);
}
