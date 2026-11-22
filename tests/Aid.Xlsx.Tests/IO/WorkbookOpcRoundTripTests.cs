using System.Text;
using System.Xml.Linq;
using Aid.Xlsx.Opc;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.IO;

[TestFixture]
public sealed class WorkbookOpcRoundTripTests
{
    private const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    [TestCase("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
    [TestCase("http://purl.oclc.org/ooxml/spreadsheetml/main", "http://purl.oclc.org/ooxml/officeDocument/relationships")]
    public void ReadsWorkbookAndStreamsWorksheetCells(string spreadsheetNamespace, string relationshipNamespace)
    {
        using var original = new MemoryStream();
        using (var workbook = new XlsxWorkbook())
        {
            workbook.AddWorksheet("Data");
            workbook.SaveAs(original, validate: false);
        }

        using var updated = new MemoryStream();
        using (var package = OpcPackage.Open(original))
        {
            var workbookPart = package.GetPart(package.Relationships.Single(rel => rel.Type == OfficeRelationships + "officeDocument").TargetPartName);
            var sheetRelationship = workbookPart.Relationships.Single(rel => rel.Type == OfficeRelationships + "worksheet");
            var worksheet = package.GetPart(sheetRelationship.TargetPartName);
            workbookPart.ReplaceContent(Encoding.UTF8.GetBytes($"""
                <workbook xmlns="{spreadsheetNamespace}" xmlns:r="{relationshipNamespace}">
                  <fileSharing readOnlyRecommended="1" userName="Reader"/>
                  <workbookPr date1904="true"/>
                  <sheets><sheet name="Data" sheetId="1" r:id="{sheetRelationship.Id}"/></sheets>
                  <definedNames><definedName name="Source" hidden="1" comment="Input cell">Data!$A$1</definedName></definedNames>
                  <calcPr calcMode="manual" refMode="R1C1"/>
                </workbook>
                """));
            worksheet.ReplaceContent(Encoding.UTF8.GetBytes($"""
                <worksheet xmlns="{spreadsheetNamespace}">
                  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>abc</t></is></c></row></sheetData>
                  <mergeCells><mergeCell ref="A1:B1"/></mergeCells>
                  <pageMargins left="0.4" right="0.5" top="0.6" bottom="0.7" header="0.2" footer="0.3"/>
                </worksheet>
                """));
            package.Save(updated);
        }

        using var loaded = new XlsxWorkbook(updated);
        Assert.That(loaded.Use1904DateSystem, Is.True);
        Assert.That(loaded.FileSharing.ReadOnlyRecommended, Is.True);
        Assert.That(loaded.FileSharing.UserName, Is.EqualTo("Reader"));
        Assert.That(loaded.CalculateMode, Is.EqualTo(XlsxCalculateMode.Manual));
        Assert.That(loaded.ReferenceStyle, Is.EqualTo(XlsxReferenceStyle.R1C1));
        var name = loaded.DefinedNames.Single();
        Assert.That(name.RefersTo, Is.EqualTo("Data!$A$1"));
        Assert.That(name.Comment, Is.EqualTo("Input cell"));
        Assert.That(name.Visible, Is.False);
        var sheet = loaded.Worksheet("Data");
        Assert.That(sheet.Cell("A1").GetString(), Is.EqualTo("abc"));
        Assert.That(sheet.MergedRanges.Select(range => range.RangeAddress.ToString()), Is.EqualTo(new[] { "A1:B1" }));
        Assert.That(sheet.PageSetup.Margins.Left, Is.EqualTo(0.4));
    }

    [Test]
    public void NewThemeRetainsCustomRgbAndSystemFallbackColors()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XlsxWorkbook())
        {
            workbook.AddWorksheet("Colors");
            workbook.Theme.Accent1 = XlsxColor.FromHexRgb("123456");
            workbook.Theme.Text1 = XlsxColor.FromHexRgb("234567");
            workbook.SaveAs(stream, validate: true);
        }

        stream.Position = 0;
        using var loaded = new XlsxWorkbook(stream);
        Assert.That(loaded.Theme.Accent1, Is.EqualTo(XlsxColor.FromHexRgb("123456")));
        Assert.That(loaded.Theme.Text1, Is.EqualTo(XlsxColor.FromHexRgb("234567")));
    }

    [Test]
    public void EditingWorkbookKeepsUnknownPartsAndTheirRelationships()
    {
        using var original = new MemoryStream();
        using (var workbook = new XlsxWorkbook())
        {
            workbook.AddWorksheet("Sheet").Cell("A1").Value = 1;
            workbook.SaveAs(original, validate: false);
        }

        using var withExtras = new MemoryStream();
        var payload = new byte[] { 0, 1, 200, 255 };
        using (var opc = OpcPackage.Open(original))
        {
            var workbookPartName = opc.Relationships.Single(rel => rel.Type == OfficeRelationships + "officeDocument").TargetPartName;
            var workbookPart = opc.GetPart(workbookPartName);
            var extras = opc.AddPart("/test/extra.dat", "application/octet-stream");
            extras.ReplaceContent(payload);
            extras.Relationships.Add("outside", "https://example.invalid/link", "https://example.invalid/test#section", isExternal: true);
            workbookPart.Relationships.Add("userPart", "https://example.invalid/extra-part", "/test/extra.dat");
            opc.AddPart("/test/standalone.xml", "application/xml").ReplaceContent("<standalone/>"u8);
            opc.Save(withExtras);
        }

        using (var workbook = new XlsxWorkbook(withExtras))
        {
            workbook.Worksheet("Sheet").Cell("A1").Value = 2;
            workbook.Save(new Workbooks.SaveOptions { ValidatePackage = false });
        }

        using var edited = new XlsxWorkbook(withExtras);
        Assert.That(edited.Worksheet("Sheet").Cell("A1").GetDouble(), Is.EqualTo(2));
        using var saved = OpcPackage.Open(withExtras);
        var usedWorkbookPart = saved.GetPart(saved.Relationships.Single(rel => rel.Type == OfficeRelationships + "officeDocument").TargetPartName);
        Assert.That(usedWorkbookPart.Relationships["userPart"].Target, Is.EqualTo("/test/extra.dat"));
        var savedExtras = saved.GetPart("/test/extra.dat");
        Assert.That(savedExtras.Relationships["outside"].Target, Is.EqualTo("https://example.invalid/test#section"));
        using var raw = savedExtras.OpenRead();
        var actual = new byte[payload.Length];
        raw.ReadExactly(actual);
        Assert.That(actual, Is.EqualTo(payload));
        Assert.That(raw.ReadByte(), Is.EqualTo(-1));
        Assert.That(saved.Parts.ContainsKey("/test/standalone.xml"), Is.True);
    }

    [Test]
    public void EditingMacroEnabledWorkbookRetainsVbaProjectBytes()
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LO\xlsm\vba-user-function.xlsm"));
        using var target = new MemoryStream();
        var before = ReadVbaProject(source);

        using (var workbook = new XlsxWorkbook(source))
        {
            workbook.Worksheet(1).Cell("Z1").Value = "Updated";
            workbook.SaveAs(target, validate: false);
        }

        using var edited = new XlsxWorkbook(target);
        Assert.That(edited.Worksheet(1).Cell("Z1").GetString(), Is.EqualTo("Updated"));
        Assert.That(ReadVbaProject(target), Is.EqualTo(before));
    }

    [Test]
    public void ConvertsWorkbookPackageToTemplateWhenSavingUnderANewExtension()
    {
        var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var source = Path.Combine(directory, "original.xlsx");
            var template = Path.Combine(directory, "template.xltx");
            using var workbook = new XlsxWorkbook();
            workbook.AddWorksheet("Data").Cell("A1").Value = 42;
            workbook.SaveAs(source, validate: false);

            workbook.SaveAs(template, validate: false);

            using var file = File.OpenRead(template);
            using var opc = OpcPackage.Open(file);
            var workbookPart = opc.GetPart(opc.Relationships.Single(rel => rel.Type == OfficeRelationships + "officeDocument").TargetPartName);
            Assert.That(workbookPart.ContentType, Is.EqualTo("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"));
            using var reloaded = new XlsxWorkbook(file);
            Assert.That(reloaded.Worksheet("Data").Cell("A1").GetDouble(), Is.EqualTo(42));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Test]
    public void EditingDocumentPropertiesPreservesUnmodeledMetadata()
    {
        const string coreRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        XNamespace core = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        XNamespace extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        XNamespace variant = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        using var original = new MemoryStream();
        using (var workbook = new XlsxWorkbook())
        {
            workbook.AddWorksheet("Data");
            workbook.Properties.Company = "Example Company";
            workbook.Properties.Manager = "Old Manager";
            workbook.Properties.Title = "Original title";
            workbook.CustomProperties.Add("Remove", false);
            workbook.CustomProperties.Add("Edit", 1.25);
            workbook.SaveAs(original, validate: true);
        }

        using var modified = new MemoryStream();
        using (var package = OpcPackage.Open(original))
        {
            var coreLink = package.Relationships.Single(rel => rel.Type == coreRelationship);
            var originalCore = package.GetPart(coreLink.TargetPartName);
            var coreXml = ReadXml(originalCore);
            coreXml.Root.SetElementValue(core + "revision", "7");
            var movedCore = package.AddPart("/metadata/workbook.xml", originalCore.ContentType);
            ReplaceXml(movedCore, coreXml);
            package.RemovePart(originalCore.PartName);
            package.Relationships.Add(coreLink.Id, coreLink.Type, "metadata/workbook.xml");

            var customPart = package.GetPart("/docProps/custom.xml");
            var customXml = ReadXml(customPart);
            var data = new XElement(custom + "property",
                new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                new XAttribute("name", "Integer"), new XAttribute("pid", 4),
                new XElement(variant + "i4", 42));
            customXml.Root.Add(data);
            ReplaceXml(customPart, customXml);

            var extendedPart = package.GetPart("/docProps/app.xml");
            var extendedXml = ReadXml(extendedPart);
            extendedXml.Root.SetElementValue(extended + "Application", "Another editor");
            ReplaceXml(extendedPart, extendedXml);
            package.Save(modified);
        }

        using (var workbook = new XlsxWorkbook(modified))
        {
            Assert.That(workbook.Properties.Company, Is.EqualTo("Example Company"));
            Assert.That(workbook.Properties.Title, Is.EqualTo("Original title"));
            workbook.Properties.Title = null;
            Assert.That(workbook.CustomProperties.CustomProperty("Edit").GetValue<double>(), Is.EqualTo(1.25));
            workbook.Properties.Manager = "New Manager";
            workbook.CustomProperties.Delete("Remove");
            workbook.CustomProperties.CustomProperty("Edit").Value = 2.5;
            workbook.CustomProperties.Add("New", true);
            workbook.Save(new Workbooks.SaveOptions { ValidatePackage = true });
        }

        using var result = OpcPackage.Open(modified);
        var savedCore = result.Relationships.Single(rel => rel.Type == coreRelationship);
        Assert.That(savedCore.TargetPartName, Is.EqualTo("/metadata/workbook.xml"));
        var coreProperties = ReadXml(result.GetPart(savedCore.TargetPartName)).Root;
        Assert.That((string)coreProperties.Element(core + "revision"), Is.EqualTo("7"));
        Assert.That(coreProperties.Elements().Any(e => e.Name.LocalName == "title"), Is.False);
        var customProperties = ReadXml(result.GetPart("/docProps/custom.xml")).Root.Elements().ToArray();
        var integer = customProperties.Single(e => (string)e.Attribute("name") == "Integer");
        Assert.That((int)integer.Element(variant + "i4"), Is.EqualTo(42));
        Assert.That(customProperties.Select(e => (int)e.Attribute("pid")), Is.Unique);
        Assert.That(customProperties.Select(e => (string)e.Attribute("name")), Is.EquivalentTo(new[] { "Integer", "Edit", "New" }));
        var extendedProperties = ReadXml(result.GetPart("/docProps/app.xml")).Root;
        Assert.That((string)extendedProperties.Element(extended + "Application"), Is.EqualTo("Another editor"));
        Assert.That((string)extendedProperties.Element(extended + "Manager"), Is.EqualTo("New Manager"));
    }

    private static XDocument ReadXml(OpcPart part)
    {
        using var stream = part.OpenRead();
        return XDocument.Load(stream);
    }

    private static void ReplaceXml(OpcPart part, XDocument xml)
    {
        using var stream = new MemoryStream();
        xml.Save(stream);
        part.ReplaceContent(stream.ToArray());
    }

    private static byte[] ReadVbaProject(Stream stream)
    {
        using var opc = OpcPackage.Open(stream);
        var workbookPart = opc.GetPart(opc.Relationships.Single(rel => rel.Type == OfficeRelationships + "officeDocument").TargetPartName);
        var vbaPart = opc.GetPart(workbookPart.Relationships.Single(rel => rel.Type.EndsWith("/vbaProject", StringComparison.Ordinal)).TargetPartName);
        using var raw = vbaPart.OpenRead();
        using var copy = new MemoryStream();
        raw.CopyTo(copy);
        return copy.ToArray();
    }
}
