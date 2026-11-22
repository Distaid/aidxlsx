using System.IO.Compression;
using System.Text;
using NUnit.Framework;
using Aid.Xlsx.Opc.Packaging;

namespace Aid.Xlsx.Tests.Opc.Packaging;

[TestFixture]
public class OpcPackageTests
{
    private const string Types = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string RelType = "https://example.com/types/test-part";

    [Test]
    public void SavesDeterministicallyAndRetainsUnchangedPartsAndTargets()
    {
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        var macro = new byte[] { 0, 1, 255, 128 };
        using (var created = OpcPackage.Create())
        {
            var workbook = created.AddPart("/xl/workbook.xml", "application/xml");
            workbook.ReplaceContent("<workbook/>"u8);
            created.AddPart("/xl/vbaProject.bin", "application/vnd.ms-office.vbaProject").ReplaceContent(macro);
            created.AddPart("/customXml/item%20one.xml", "application/xml").ReplaceContent("<unknown/>"u8);
            created.Relationships.Add("rId1", RelType, "/xl/workbook.xml");
            workbook.Relationships.Add("rId1", RelType, "vbaProject.bin");
            workbook.Relationships.Add("rId2", RelType, "../customXml/item%20one.xml#name");
            workbook.Relationships.Add("rId3", RelType, "https://example.com/data?a=1&b=2#link", isExternal: true);
            created.Save(first);
        }

        first.Position = 0;
        using (var loaded = OpcPackage.Open(first))
        {
            var workbook = loaded.GetPart("/XL/WORKBOOK.XML");
            Assert.That(loaded.Relationships["rId1"].TargetPartName, Is.EqualTo("/xl/workbook.xml"));
            Assert.That(workbook.Relationships["rId1"].TargetPartName, Is.EqualTo("/xl/vbaProject.bin"));
            Assert.That(workbook.Relationships["rId2"].TargetPartName, Is.EqualTo("/customXml/item%20one.xml"));
            Assert.That(workbook.Relationships["rId2"].Target, Is.EqualTo("../customXml/item%20one.xml#name"));
            Assert.That(workbook.Relationships["rId3"].Target, Is.EqualTo("https://example.com/data?a=1&b=2#link"));
            Assert.That(workbook.Relationships["rId3"].TargetPartName, Is.Null);
            using var macroStream = loaded.GetPart("/xl/vbaProject.bin").OpenRead();
            using var contents = new MemoryStream();
            macroStream.CopyTo(contents);
            Assert.That(contents.ToArray(), Is.EqualTo(macro));
            loaded.Save(second);
        }

        Assert.That(first.ToArray(), Is.EqualTo(second.ToArray()));
        second.Position = 0;
        using var archive = new ZipArchive(second, ZipArchiveMode.Read);
        Assert.That(archive.Entries.Select(entry => entry.FullName), Is.Ordered.Using<string>(StringComparer.Ordinal));
        Assert.That(archive.Entries.All(entry => entry.LastWriteTime.Year == 1980), Is.True);
    }

    [Test]
    public void ReadsDefaultContentTypesAndRemovesIncomingRelationships()
    {
        using var zip = MakeZip($"<Types xmlns='{Types}'><Default Extension='XML' ContentType='application/xml'/></Types>",
            ("source.xml", "<source/>"), ("target.xml", "<target/>"));
        using var package = OpcPackage.Open(zip);
        Assert.That(package.GetPart("/source.xml").ContentType, Is.EqualTo("application/xml"));
        var relationships = package.GetPart("/source.xml").Relationships;
        relationships.Add("rId1", RelType, "target.xml#fragment");
        relationships.Add("rId2", RelType, "target.xml", isExternal: true);
        package.Relationships.Add("rId1", RelType, "target.xml");
        Assert.That(package.RemovePart("/TARGET.XML"), Is.True);
        Assert.That(package.Relationships, Is.Empty);
        Assert.That(relationships.Select(link => link.Id), Is.EqualTo(new[] { "rId2" }));
    }

    [Test]
    public void ReplacesLazyPartContentsWhenTheWriteStreamCloses()
    {
        using var zip = MakeZip($"<Types xmlns='{Types}'><Default Extension='xml' ContentType='application/xml'/></Types>",
            ("source.xml", "<old-content/>"));
        using var package = OpcPackage.Open(zip);
        var part = package.GetPart("/source.xml");
        using (var writer = part.OpenWrite())
        {
            writer.Write("<new/>"u8);
            Assert.That(part.Length, Is.EqualTo("<old-content/>".Length));
        }

        Assert.That(part.Length, Is.EqualTo("<new/>".Length));
        using var saved = new MemoryStream();
        package.Save(saved);
        saved.Position = 0;
        using var reopened = OpcPackage.Open(saved);
        using var reader = new StreamReader(reopened.GetPart("/source.xml").OpenRead());
        Assert.That(reader.ReadToEnd(), Is.EqualTo("<new/>"));
    }

    [TestCase("../payload.xml")]
    [TestCase("xl/%2fpart.xml")]
    [TestCase("xl/%2e%2e/payload.xml")]
    [TestCase("xl\\payload.xml")]
    [TestCase("/payload.xml")]
    public void RejectsInvalidPartNames(string name)
    {
        using var zip = MakeZip($"<Types xmlns='{Types}'><Default Extension='xml' ContentType='application/xml'/></Types>", (name, "text"));
        Assert.That(() => OpcPackage.Open(zip), Throws.TypeOf<InvalidDataException>().With.Message.Contains("part name"));
    }

    [Test]
    public void AppliesLimitsAndRejectsDuplicateEntries()
    {
        using var zip = MakeZip($"<Types xmlns='{Types}'/>", ("a.xml", "text"), ("A.xml", "text"));
        Assert.That(() => OpcPackage.Open(zip, new OpcOpenOptions { MaxEntries = 1 }), Throws.TypeOf<InvalidDataException>().With.Message.Contains("too many ZIP entries"));
        Assert.That(() => OpcPackage.Open(zip, new OpcOpenOptions { MaxUncompressedBytes = 2 }), Throws.TypeOf<InvalidDataException>().With.Message.Contains("size limit"));
        Assert.That(() => OpcPackage.Open(zip), Throws.TypeOf<InvalidDataException>().With.Message.Contains("Duplicate"));

        using var compressed = MakeZip($"<Types xmlns='{Types}'/>", ("a.xml", new string('a', 4096)));
        Assert.That(() => OpcPackage.Open(compressed, new OpcOpenOptions { MaxCompressionRatio = 10 }), Throws.TypeOf<InvalidDataException>().With.Message.Contains("compression ratio"));
    }

    [TestCase("<!DOCTYPE Types SYSTEM 'http://example.com/metadata.dtd'>")]
    [TestCase("<Types/>")]
    public void RejectsMalformedMetadata(string prefix)
    {
        using var zip = MakeZip(prefix + $"<Types xmlns='{Types}'/>");
        Assert.That(() => OpcPackage.Open(zip), Throws.TypeOf<InvalidDataException>().With.Message.Contains("[Content_Types].xml"));
    }

    [Test]
    public void RejectsRelationshipThatEscapesPackageRoot()
    {
        using var package = OpcPackage.Create();
        var part = package.AddPart("/xl/workbook.xml", "application/xml");
        Assert.That(() => part.Relationships.Add("rId1", RelType, "../../other.xml"), Throws.TypeOf<InvalidDataException>());
    }

    private static MemoryStream MakeZip(string contentTypes, params (string Name, string Content)[] entries)
    {
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write("[Content_Types].xml", contentTypes);
            foreach (var (name, content) in entries)
                Write(name, content);

            void Write(string name, string content)
            {
                using var output = zip.CreateEntry(name).Open();
                output.Write(Encoding.UTF8.GetBytes(content));
            }
        }

        stream.Position = 0;
        return stream;
    }
}
