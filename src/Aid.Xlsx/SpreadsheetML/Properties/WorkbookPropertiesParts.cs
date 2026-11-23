using System.Xml.Linq;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Opc.Packaging;

namespace Aid.Xlsx.SpreadsheetML.Properties;

internal static class WorkbookPropertiesParts
{
    private const string CoreRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    private const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string CustomRelationship = OfficeRelationships + "custom-properties";
    private const string ExtendedRelationship = OfficeRelationships + "extended-properties";

    internal static void Read(OpcPackage package, XlsxWorkbook workbook)
    {
        using (var core = GetPart(package, CoreRelationship)?.OpenRead())
            if (core is not null)
                CorePropertiesXml.Read(core, workbook.Properties);

        using (var extended = GetPart(package, ExtendedRelationship)?.OpenRead())
            if (extended is not null)
                ExtendedPropertiesXml.Read(extended, workbook);

        using var custom = GetPart(package, CustomRelationship)?.OpenRead();
        if (custom is not null)
            CustomPropertiesXml.Read(custom, workbook.CustomProperties);
    }

    internal static bool HasExtendedProperties(OpcPackage package) => GetPart(package, ExtendedRelationship) is not null;

    internal static void Write(OpcPackage package, XlsxWorkbook workbook, string? newExtendedRelationshipId)
    {
        var corePart = GetPart(package, CoreRelationship);
        XDocument core;
        using (var existing = corePart?.OpenRead())
            core = CorePropertiesXml.Update(existing, workbook.Properties);
        corePart ??= AddPart(package, CoreRelationship, "/package/services/metadata/core-properties/core.psmdcp",
            "application/vnd.openxmlformats-package.core-properties+xml");
        ReplaceXml(corePart, core);

        var extendedPart = GetPart(package, ExtendedRelationship);
        XDocument extended;
        using (var existing = extendedPart?.OpenRead())
            extended = ExtendedPropertiesXml.Update(existing, workbook);
        extendedPart ??= AddPart(package, ExtendedRelationship, "/docProps/app.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml", newExtendedRelationshipId);
        ReplaceXml(extendedPart, extended);

        var customPart = GetPart(package, CustomRelationship);
        XDocument? custom;
        using (var existing = customPart?.OpenRead())
            custom = CustomPropertiesXml.Update(existing, workbook.CustomProperties);
        if (custom is not null)
        {
            customPart ??= AddPart(package, CustomRelationship, "/docProps/custom.xml",
                "application/vnd.openxmlformats-officedocument.custom-properties+xml");
            ReplaceXml(customPart, custom);
        }
        else if (customPart is not null)
        {
            package.RemovePart(customPart.PartName);
        }
    }

    private static OpcPart? GetPart(OpcPackage package, string type)
    {
        var partName = package.Relationships.FirstOrDefault(rel => rel.Type == type && !rel.IsExternal)?.TargetPartName;
        return partName is not null && package.TryGetPart(partName, out var part) ? part : null;
    }

    private static OpcPart AddPart(OpcPackage package, string relationType, string defaultName, string contentType,
        string? preferredId = null)
    {
        var name = defaultName;
        for (var suffix = 2; package.Parts.ContainsKey(name); suffix++)
            name = defaultName.Insert(defaultName.LastIndexOf('.'), suffix.ToString(System.Globalization.CultureInfo.InvariantCulture));

        var relationship = package.Relationships.FirstOrDefault(rel => rel.Type == relationType && !rel.IsExternal);
        var id = relationship?.Id;
        if (id is null)
        {
            var usedIds = package.Relationships.Select(rel => rel.Id).ToHashSet(StringComparer.Ordinal);
            if (preferredId is not null && !usedIds.Contains(preferredId))
                id = preferredId;
            else
            {
                for (var index = 1; !usedIds.Add(id = FormattableString.Invariant($"rId{index}")); index++)
                {
                }
            }
        }
        else
        {
            package.Relationships.Remove(id);
        }

        var part = package.AddPart(name, contentType);
        package.Relationships.Add(id, relationType, name[1..]);
        return part;
    }

    private static void ReplaceXml(OpcPart part, XDocument xml)
    {
        using var stream = new MemoryStream();
        xml.Save(stream, System.Xml.Linq.SaveOptions.DisableFormatting);
        part.ReplaceContent(stream.GetBuffer().AsSpan(0, (int)stream.Length));
    }
}
