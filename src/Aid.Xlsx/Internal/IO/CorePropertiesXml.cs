using System.Globalization;
using System.Xml.Linq;

namespace Aid.Xlsx.Internal.IO;

internal static class CorePropertiesXml
{
    private static readonly XNamespace Core = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DublinCore = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Terms = "http://purl.org/dc/terms/";
    private static readonly XNamespace SchemaInstance = "http://www.w3.org/2001/XMLSchema-instance";

    internal static void Read(Stream stream, XLWorkbookProperties properties)
    {
        var root = XDocumentExtensions.Load(stream)?.Root;
        if (root?.Name != Core + "coreProperties")
            return;

        properties.Author = (string?)root.Element(DublinCore + "creator");
        properties.Category = (string?)root.Element(Core + "category");
        properties.Comments = (string?)root.Element(DublinCore + "description");
        properties.Keywords = (string?)root.Element(Core + "keywords");
        properties.LastModifiedBy = (string?)root.Element(Core + "lastModifiedBy");
        properties.Status = (string?)root.Element(Core + "contentStatus");
        properties.Subject = (string?)root.Element(DublinCore + "subject");
        properties.Title = (string?)root.Element(DublinCore + "title");
        if (TryReadDate(root, "created", out var created))
            properties.Created = created;
        if (TryReadDate(root, "modified", out var modified))
            properties.Modified = modified;
    }

    internal static XDocument Update(Stream? existing, XLWorkbookProperties properties)
    {
        var document = existing is null ? null : XDocumentExtensions.Load(existing);
        if (document?.Root is not { } root || root.Name != Core + "coreProperties")
        {
            root = new XElement(Core + "coreProperties");
            document = new XDocument(root);
        }

        root.SetAttributeValue(XNamespace.Xmlns + "dc", DublinCore.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "dcterms", Terms.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "xsi", SchemaInstance.NamespaceName);
        root.SetElementValue(DublinCore + "creator", properties.Author);
        root.SetElementValue(DublinCore + "title", properties.Title);
        root.SetElementValue(DublinCore + "subject", properties.Subject);
        root.SetElementValue(Core + "category", properties.Category);
        root.SetElementValue(Core + "keywords", properties.Keywords);
        root.SetElementValue(DublinCore + "description", properties.Comments);
        root.SetElementValue(Core + "contentStatus", properties.Status);
        root.SetElementValue(Core + "lastModifiedBy", properties.LastModifiedBy);
        WriteDate(root, "created", properties.Created);
        WriteDate(root, "modified", properties.Modified);
        return document;
    }

    private static bool TryReadDate(XElement root, string name, out DateTime value) =>
        DateTime.TryParse((string?)root.Element(Terms + name), CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out value);

    private static void WriteDate(XElement root, string name, DateTime value)
    {
        var utc = value == DateTime.MinValue ? DateTime.UtcNow : value.Kind == DateTimeKind.Local
            ? value.ToUniversalTime() : DateTime.SpecifyKind(value, DateTimeKind.Utc);
        var elementName = Terms + name;
        root.SetElementValue(elementName, utc.ToString("O", CultureInfo.InvariantCulture));
        root.Element(elementName)!.SetAttributeValue(SchemaInstance + "type", "dcterms:W3CDTF");
    }
}
