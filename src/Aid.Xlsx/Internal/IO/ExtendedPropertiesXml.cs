using System.Xml.Linq;

namespace Aid.Xlsx.Internal.IO;

internal static class ExtendedPropertiesXml
{
    private static readonly XNamespace Extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace Variant = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    internal static void Read(Stream stream, XLWorkbook workbook)
    {
        var root = XDocumentExtensions.Load(stream)?.Root;
        if (root is null)
            return;

        workbook.Properties.Company = (string?)root.Element(Extended + "Company");
        workbook.Properties.Manager = (string?)root.Element(Extended + "Manager");
    }

    internal static XDocument Update(Stream? existing, XLWorkbook workbook)
    {
        var root = existing is null ? null : XDocumentExtensions.Load(existing)?.Root;
        if (root?.Name != Extended + "Properties")
            root = new XElement(Extended + "Properties");

        if ((string?)root.Attribute("xmlns") == Extended.NamespaceName)
            root.Attribute("xmlns")!.Remove();
        root.SetAttributeValue(XNamespace.Xmlns + "ap", Extended.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "vt", Variant.NamespaceName);
        AddDefault(root, "Application", "Microsoft Excel");
        AddDefault(root, "DocSecurity", "0");
        AddDefault(root, "ScaleCrop", "false");

        var worksheetNames = workbook.WorksheetsInternal.Cast<XLWorksheet>()
            .OrderBy(sheet => sheet.Position).Select(sheet => sheet.Name).ToList();
        var namedRanges = GetRangeNames(workbook);
        SetVector(root, "TitlesOfParts", "lpstr", worksheetNames.Concat(namedRanges)
            .Select(name => new XElement(Variant + "lpstr", name)).ToList());
        SetVector(root, "HeadingPairs", "variant",
        [
            new XElement(Variant + "variant", new XElement(Variant + "lpstr", "Worksheets")),
            new XElement(Variant + "variant", new XElement(Variant + "i4", worksheetNames.Count)),
            new XElement(Variant + "variant", new XElement(Variant + "lpstr", "Named Ranges")),
            new XElement(Variant + "variant", new XElement(Variant + "i4", namedRanges.Count))
        ]);
        SetOptional(root, "Manager", workbook.Properties.Manager);
        SetOptional(root, "Company", workbook.Properties.Company);
        return root.Document ?? new XDocument(root);
    }

    private static void AddDefault(XElement root, string name, string value)
    {
        if (root.Element(Extended + name) is null)
            root.SetElementValue(Extended + name, value);
    }

    private static void SetVector(XElement root, string name, string kind, IReadOnlyCollection<XElement> entries)
    {
        var vector = new XElement(Extended + name, new XElement(Variant + "vector",
            new XAttribute("baseType", kind), new XAttribute("size", entries.Count), entries));
        var existing = root.Element(vector.Name);
        if (existing is null)
            root.Element(Extended + "DocSecurity")!.AddBeforeSelf(vector);
        else
            existing.ReplaceWith(vector);
    }

    private static void SetOptional(XElement root, string name, string? value)
    {
        if (value is not null)
            root.SetElementValue(Extended + name, string.IsNullOrWhiteSpace(value) ? null : value);
    }

    private static List<string> GetRangeNames(XLWorkbook workbook)
    {
        var names = new List<string>();
        foreach (var sheet in workbook.WorksheetsInternal)
        {
            names.AddRange(sheet.DefinedNames.Select<XLDefinedName, string>(namedRange => $"{sheet.Name}!{namedRange.Name}"));
            names.Add($"{sheet.Name}!Print_Area");
            names.Add($"{sheet.Name}!Print_Titles");
        }

        names.AddRange(workbook.DefinedNamesInternal.Select<XLDefinedName, string>(namedRange => namedRange.Name));
        return names;
    }
}
