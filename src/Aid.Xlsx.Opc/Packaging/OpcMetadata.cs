using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Aid.Xlsx.Opc.Packaging;

internal static class OpcMetadata
{
    internal const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal static void ReadContentTypes(ZipArchiveEntry entry, long maxCharacters,
        Dictionary<string, string> defaults, Dictionary<string, string> overrides)
    {
        ReadElements(entry, maxCharacters, "Types", ContentTypesNamespace, reader =>
        {
            var contentType = RequiredAttribute(reader, "ContentType");
            ValidateContentType(contentType);
            switch (reader.LocalName)
            {
                case "Default":
                    var extension = RequiredAttribute(reader, "Extension");
                    if (extension.Any(c => c is '.' or '/' or '\\' || char.IsWhiteSpace(c)))
                        throw new InvalidDataException($"Invalid OPC extension '{extension}'.");
                    AddUnique(defaults, extension, contentType);
                    break;
                case "Override":
                    var partName = OpcPartName.Validate(RequiredAttribute(reader, "PartName"));
                    AddUnique(overrides, partName, contentType);
                    break;
                default:
                    throw new InvalidDataException($"Unexpected OPC content type element '{reader.LocalName}'.");
            }
        });
    }

    internal static void ReadRelationships(ZipArchiveEntry entry, long maxCharacters, OpcRelationshipCollection relationships)
    {
        ReadElements(entry, maxCharacters, "Relationships", RelationshipsNamespace, reader =>
        {
            if (reader.LocalName != "Relationship")
                throw new InvalidDataException($"Unexpected OPC relationship element '{reader.LocalName}'.");

            var isExternal = reader.GetAttribute("TargetMode") switch
            {
                null or "Internal" => false,
                "External" => true,
                var mode => throw new InvalidDataException($"Invalid OPC target mode '{mode}'."),
            };

            relationships.Add(RequiredAttribute(reader, "Id"), RequiredAttribute(reader, "Type"),
                RequiredAttribute(reader, "Target"), isExternal);
        });
    }

    internal static void WriteContentTypes(Stream stream, Dictionary<string, string> defaults, Dictionary<string, string> overrides)
    {
        using XmlWriter writer = CreateWriter(stream);
        writer.WriteStartDocument();
        writer.WriteStartElement("Types", ContentTypesNamespace);
        foreach ((string? extension, string? contentType) in defaults.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            writer.WriteStartElement("Default", ContentTypesNamespace);
            writer.WriteAttributeString("Extension", extension);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        foreach ((string? partName, string? contentType) in overrides.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            writer.WriteStartElement("Override", ContentTypesNamespace);
            writer.WriteAttributeString("PartName", partName);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        writer.WriteEndDocument();
    }

    internal static void WriteRelationships(Stream stream, OpcRelationshipCollection relationships)
    {
        using XmlWriter writer = CreateWriter(stream);
        writer.WriteStartDocument();
        writer.WriteStartElement("Relationships", RelationshipsNamespace);
        foreach (OpcRelationship? relationship in relationships.OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            writer.WriteStartElement("Relationship", RelationshipsNamespace);
            writer.WriteAttributeString("Id", relationship.Id);
            writer.WriteAttributeString("Type", relationship.Type);
            writer.WriteAttributeString("Target", relationship.Target);
            if (relationship.IsExternal)
                writer.WriteAttributeString("TargetMode", "External");
            writer.WriteEndElement();
        }

        writer.WriteEndDocument();
    }

    internal static void ValidateContentType(string contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType) || !contentType.Contains('/') || contentType.Any(char.IsControl))
            throw new InvalidDataException($"Invalid OPC content type '{contentType}'.");
    }

    private static void ReadElements(ZipArchiveEntry entry, long maxCharacters, string root, string ns, Action<XmlReader> readElement)
    {
        try
        {
            using Stream stream = entry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = maxCharacters,
            });

            reader.MoveToContent();
            var emptyRoot = reader.IsEmptyElement;
            reader.ReadStartElement(root, ns);
            if (!emptyRoot)
            {
                while (reader.MoveToContent() == XmlNodeType.Element)
                {
                    if (reader.NamespaceURI != ns)
                        throw new InvalidDataException($"Unexpected OPC metadata namespace '{reader.NamespaceURI}'.");

                    readElement(reader);
                    var emptyElement = reader.IsEmptyElement;
                    reader.ReadStartElement();
                    if (!emptyElement)
                        reader.ReadEndElement();
                }

                reader.ReadEndElement();
            }

            if (reader.MoveToContent() != XmlNodeType.None)
                throw new InvalidDataException("Unexpected content after the OPC metadata root.");
        }
        catch (Exception exception) when (exception is XmlException or InvalidDataException)
        {
            throw new InvalidDataException($"Invalid OPC metadata in '{entry.FullName}'. {exception.Message}", exception);
        }
    }

    private static XmlWriter CreateWriter(Stream stream) => XmlWriter.Create(stream, new XmlWriterSettings
    {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        NewLineHandling = NewLineHandling.Entitize,
    });

    private static string RequiredAttribute(XmlReader reader, string attribute)
    {
        var value = reader.GetAttribute(attribute);
        return string.IsNullOrEmpty(value)
            ? throw new InvalidDataException($"Missing '{attribute}' on OPC element '{reader.LocalName}'.")
            : value;
    }

    private static void AddUnique(Dictionary<string, string> types, string name, string contentType)
    {
        if (!types.TryAdd(name, contentType))
            throw new InvalidDataException($"Duplicate OPC content type declaration '{name}'.");
    }
}
