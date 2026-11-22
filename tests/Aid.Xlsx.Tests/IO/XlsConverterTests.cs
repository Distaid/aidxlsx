using System;
using System.IO;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.IO;

[TestFixture]
public class XlsConverterTests
{
    [Test]
    public void ConvertsWorkbookWrittenByAnIndependentBiff8Writer()
    {
        using var source = OpenFixture();
        using var destination = new MemoryStream();

        XlsConverter.Convert(source, destination);

        Assert.That(source.CanRead, Is.True);
        Assert.That(destination.CanWrite, Is.True);
        destination.Position = 0;
        using var workbook = new XLWorkbook(destination);
        var sheet = workbook.Worksheet("Résumé");
        Assert.Multiple(() =>
        {
            Assert.That(workbook.Use1904DateSystem, Is.True);
            Assert.That(sheet.Cell("A1").GetString(), Is.EqualTo("Καλημέρα"));
            Assert.That(sheet.Cell("B1").GetDouble(), Is.EqualTo(-12.5));
            Assert.That(sheet.Cell("B1").Style.NumberFormat.Format, Is.EqualTo("0.0000"));
            Assert.That(sheet.Cell("A2").GetBoolean(), Is.True);
            Assert.That(sheet.Cell("B2").GetDateTime(), Is.EqualTo(new DateTime(2020, 2, 29)));
            Assert.That(sheet.Cell("A3").GetString(), Is.EqualTo("Merged"));
            Assert.That(sheet.Cell("B4").IsMerged(), Is.True);
            Assert.That(sheet.Cell("A5").GetString(), Is.EqualTo(new string('水', 5000)));
            Assert.That(sheet.Cell("A6").GetString(), Is.EqualTo(new string('z', 9000)));
            Assert.That(sheet.Cell("A7").GetString(), Is.Empty);
            Assert.That(sheet.Row(8).Height, Is.EqualTo(22.5));
            Assert.That(sheet.Column(3).IsHidden, Is.True);
            Assert.That(sheet.Column(3).Width, Is.EqualTo(22).Within(0.01));
            Assert.That(workbook.Worksheet("Hidden").Visibility, Is.EqualTo(XLWorksheetVisibility.Hidden));
            Assert.That(workbook.Worksheet("Hidden").Cell("A1").GetDouble(), Is.EqualTo(42));
        });
    }

    [Test]
    public void UsesTheSourceStreamPosition()
    {
        using var fixture = OpenFixture();
        using var source = new MemoryStream();
        source.WriteByte(1);
        fixture.CopyTo(source);
        source.Position = 1;
        using var destination = new MemoryStream();

        XlsConverter.Convert(source, destination);

        destination.Position = 0;
        using var workbook = new XLWorkbook(destination);
        Assert.That(workbook.Worksheet("Hidden").Cell("A1").GetDouble(), Is.EqualTo(42));
    }

    [Test]
    public void LeavesDestinationIntactWhenTheInputIsTruncated()
    {
        using var source = new MemoryStream([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        using var destination = new MemoryStream();
        destination.WriteByte(42);
        destination.Position = 0;

        Assert.That(() => XlsConverter.Convert(source, destination), Throws.TypeOf<InvalidDataException>());
        Assert.That(destination.ToArray(), Is.EqualTo(new byte[] { 42 }));
    }

    // This fixture was written with xlwt 1.3.0. The long strings force BIFF CONTINUE records.
    private static Stream OpenFixture() => typeof(XlsConverterTests).Assembly
        .GetManifestResourceStream("Aid.Xlsx.Tests.IO.biff8-generated.xls");
}
