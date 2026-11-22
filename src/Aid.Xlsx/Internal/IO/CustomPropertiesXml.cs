using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace Aid.Xlsx.Internal.IO;

internal static class CustomPropertiesXml
{
    private static readonly XNamespace Custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    private static readonly XNamespace Variant = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    private const string FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    internal static void Read(Stream stream, IXLCustomProperties properties)
    {
        var root = XDocumentExtensions.Load(stream)?.Root;
        if (root is null)
            return;

        foreach (var property in root.Elements(Custom + "property"))
        {
            var name = (string?)property.Attribute("name");
            var value = property.Elements().FirstOrDefault();
            if (string.IsNullOrWhiteSpace(name) || value is null)
                continue;

            if (value.Name == Variant + "lpwstr")
                properties.Add(name, value.Value);
            else if (value.Name == Variant + "filetime")
                properties.Add(name, XmlConvert.ToDateTime(value.Value, XmlDateTimeSerializationMode.RoundtripKind));
            else if (value.Name == Variant + "r8")
                properties.Add(name, XmlConvert.ToDouble(value.Value));
            else if (value.Name == Variant + "bool")
                properties.Add(name, XmlConvert.ToBoolean(value.Value));
        }
    }

    /// <summary>Builds the updated XML, retaining variants the object model does not represent.</summary>
    /// <returns>The XML, or <see langword="null"/> if the entire part can be removed.</returns>
    internal static XDocument? Update(Stream? existing, IXLCustomProperties properties)
    {
        var root = existing is null ? null : XDocumentExtensions.Load(existing)?.Root;
        if (root?.Name != Custom + "Properties")
            root = new XElement(Custom + "Properties");

        if ((string?)root.Attribute("xmlns") == Custom.NamespaceName)
            root.Attribute("xmlns")!.Remove();
        root.SetAttributeValue(XNamespace.Xmlns + "op", Custom.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "vt", Variant.NamespaceName);

        var updated = properties.ToDictionary(property => property.Name, StringComparer.Ordinal);
        var reservedIds = new HashSet<int>();
        foreach (var element in root.Elements(Custom + "property").ToArray())
        {
            var name = (string?)element.Attribute("name");
            var variant = element.Elements().FirstOrDefault()?.Name;
            if (name is not null && updated.Remove(name, out var property))
            {
                element.ReplaceNodes(SerializeValue(property));
            }
            else if (variant == Variant + "lpwstr" || variant == Variant + "filetime" ||
                     variant == Variant + "r8" || variant == Variant + "bool")
            {
                element.Remove();
                continue;
            }

            if (int.TryParse((string?)element.Attribute("pid"), CultureInfo.InvariantCulture, out var id))
                reservedIds.Add(id);
        }

        var nextId = 2;
        foreach (var property in updated.Values)
        {
            while (reservedIds.Contains(nextId))
                nextId++;

            root.Add(new XElement(Custom + "property",
                new XAttribute("fmtid", FormatId), new XAttribute("pid", nextId++),
                new XAttribute("name", property.Name), SerializeValue(property)));
        }

        return root.HasElements ? root.Document ?? new XDocument(root) : null;
    }

    private static XElement SerializeValue(IXLCustomProperty property) => property.Type switch
    {
        XLCustomPropertyType.Text => new XElement(Variant + "lpwstr", property.GetValue<string>()),
        XLCustomPropertyType.Date => new XElement(Variant + "filetime", XmlConvert.ToString(property.GetValue<DateTime>().ToUniversalTime(), "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")),
        XLCustomPropertyType.Number => new XElement(Variant + "r8", XmlConvert.ToString(property.GetValue<double>())),
        _ => new XElement(Variant + "bool", XmlConvert.ToString(property.GetValue<bool>()))
    };
}
