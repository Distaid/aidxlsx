using System.Xml;
using static Aid.Xlsx.SpreadsheetML.Xml.OpenXmlConst;
using Aid.Xlsx.Opc.Packaging;

namespace Aid.Xlsx.SpreadsheetML.Workbooks;

internal sealed partial class WorkbookPackage
{
    internal void Validate(OpcPart workbook)
    {
        if (WorkbookPart != workbook)
            throw new InvalidDataException("The package has no relationship to its workbook part.");

        var checkedParts = new HashSet<OpcPart>();
        var pending = new Stack<(OpcPart Part, string Type)>();
        pending.Push((workbook, "officeDocument"));
        while (pending.TryPop(out (OpcPart Part, string Type) item))
        {
            if (!checkedParts.Add(item.Part))
                continue;

            ValidateXml(item.Part, item.Type);
            foreach (OpcRelationship relationship in item.Part.Relationships)
            {
                if (!TryGetOfficeType(relationship.Type, out var type) ||
                    !(type is "worksheet" or "chartsheet" or "dialogsheet" or "styles" or "sharedStrings" or "theme" or "calcChain" ||
                      IsManagedChild(relationship)))
                    continue;

                if (Resolve(relationship) is not { } target)
                    throw new InvalidDataException($"Part '{item.Part.PartName}' has a missing target for relationship '{relationship.Id}'.");
                pending.Push((target, type));
            }
        }
    }

    private void ValidateXml(OpcPart part, string type)
    {
        var root = type switch
        {
            "officeDocument" => "workbook",
            "styles" => "styleSheet",
            "sharedStrings" => "sst",
            "drawing" => "wsDr",
            "vmlDrawing" => "xml",
            "pivotTable" => "pivotTableDefinition",
            "image" => null,
            _ => type,
        };
        if (root is null)
            return;

        using Stream stream = part.OpenRead();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        try
        {
            reader.MoveToContent();
            var namespaceMatches = type switch
            {
                "drawing" => reader.NamespaceURI is "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" or "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing",
                "theme" => reader.NamespaceURI is "http://schemas.openxmlformats.org/drawingml/2006/main" or "http://purl.oclc.org/ooxml/drawingml/main",
                "vmlDrawing" => reader.NamespaceURI.Length == 0,
                _ => reader.NamespaceURI is Main2006SsNs or StrictSpreadsheetNs,
            };
            if (reader.LocalName != root || !namespaceMatches)
                throw new InvalidDataException($"Part '{part.PartName}' does not contain the expected '{root}' root element.");

            while (reader.Read())
            {
                if (type == "officeDocument" && reader.NodeType == XmlNodeType.Element && reader.Depth == 2 &&
                    reader.LocalName == "sheet" && reader.NamespaceURI is Main2006SsNs or StrictSpreadsheetNs)
                {
                    var id = reader.GetAttribute("id", RelationshipsNs) ?? reader.GetAttribute("id", StrictRelationships.TrimEnd('/'));
                    if (id is null || !part.Relationships.TryGet(id, out OpcRelationship? relationship) || Resolve(relationship) is null)
                        throw new InvalidDataException($"Part '{part.PartName}' contains a sheet without an internal relationship.");
                }
            }
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"Part '{part.PartName}' contains invalid XML.", exception);
        }
    }

    private static bool TryGetOfficeType(string qualifiedType, out string type)
    {
        if (qualifiedType.StartsWith(TransitionalRelationships, StringComparison.Ordinal))
            type = qualifiedType[TransitionalRelationships.Length..];
        else if (qualifiedType.StartsWith(StrictRelationships, StringComparison.Ordinal))
            type = qualifiedType[StrictRelationships.Length..];
        else
        {
            type = string.Empty;
            return false;
        }

        return true;
    }
}
