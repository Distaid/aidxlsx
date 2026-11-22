using System;
using System.IO;
using System.Linq;
using Aid.Xlsx.Opc;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.IO;

[TestFixture]
public sealed class WorkbookOpcRoundTripTests
{
    private const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    [Test]
    public void EditingWorkbookKeepsUnknownPartsAndTheirRelationships()
    {
        using var original = new MemoryStream();
        using (var workbook = new XLWorkbook())
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

        using (var workbook = new XLWorkbook(withExtras))
        {
            workbook.Worksheet("Sheet").Cell("A1").Value = 2;
            workbook.Save(new SaveOptions { ValidatePackage = false });
        }

        using var edited = new XLWorkbook(withExtras);
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

        using (var workbook = new XLWorkbook(source))
        {
            workbook.Worksheet(1).Cell("Z1").Value = "Updated";
            workbook.SaveAs(target, validate: false);
        }

        using var edited = new XLWorkbook(target);
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
            using var workbook = new XLWorkbook();
            workbook.AddWorksheet("Data").Cell("A1").Value = 42;
            workbook.SaveAs(source, validate: false);

            workbook.SaveAs(template, validate: false);

            using var file = File.OpenRead(template);
            using var opc = OpcPackage.Open(file);
            var workbookPart = opc.GetPart(opc.Relationships.Single(rel => rel.Type == OfficeRelationships + "officeDocument").TargetPartName);
            Assert.That(workbookPart.ContentType, Is.EqualTo("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"));
            using var reloaded = new XLWorkbook(file);
            Assert.That(reloaded.Worksheet("Data").Cell("A1").GetDouble(), Is.EqualTo(42));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
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
