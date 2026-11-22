using Aid.Xlsx.Conversion;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
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
        using var workbook = new XlsxWorkbook(destination);
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
            Assert.That(workbook.Worksheet("Hidden").Visibility, Is.EqualTo(XlsxWorksheetVisibility.Hidden));
            Assert.That(workbook.Worksheet("Hidden").Cell("A1").GetDouble(), Is.EqualTo(42));
        });
    }

    [Test]
    public void PreservesCellFormattingFromAnIndependentBiff8Writer()
    {
        // Generated with xlwt 1.3.0.
        using var source = typeof(XlsConverterTests).Assembly.GetManifestResourceStream("Aid.Xlsx.Tests.IO.biff8-styles.xls");
        using var destination = new MemoryStream();
        XlsConverter.Convert(source, destination);

        destination.Position = 0;
        using var workbook = new XlsxWorkbook(destination);
        var sheet = workbook.Worksheet(1);
        var style = sheet.Cell("A1").Style;
        Assert.Multiple(() =>
        {
            Assert.That(sheet.Cell("A1").GetDouble(), Is.EqualTo(12.5));
            Assert.That(style.Font.FontName, Is.EqualTo("Arial"));
            Assert.That(style.Font.FontSize, Is.EqualTo(14));
            Assert.That(style.Font.Bold, Is.True);
            Assert.That(style.Font.Italic, Is.True);
            Assert.That(style.Font.Underline, Is.EqualTo(XlsxFontUnderlineValues.Double));
            Assert.That(style.Font.FontColor, Is.EqualTo(XlsxColor.FromArgb(25, 50, 75)));
            Assert.That(style.Alignment.Horizontal, Is.EqualTo(XlsxAlignmentHorizontalValues.Center));
            Assert.That(style.Alignment.Vertical, Is.EqualTo(XlsxAlignmentVerticalValues.Top));
            Assert.That(style.Alignment.TextRotation, Is.EqualTo(-30));
            Assert.That(style.Alignment.WrapText, Is.True);
            Assert.That(style.Border.BottomBorder, Is.EqualTo(XlsxBorderStyleValues.Double));
            Assert.That(style.Border.BottomBorderColor, Is.EqualTo(XlsxColor.FromArgb(0, 0, 255)));
            Assert.That(style.Fill.PatternType, Is.EqualTo(XlsxFillPatternValues.Solid));
            Assert.That(style.Fill.PatternColor, Is.EqualTo(XlsxColor.FromArgb(255, 255, 0)));
            Assert.That(style.Protection.Locked, Is.False);
            Assert.That(style.Protection.Hidden, Is.True);
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
        using var workbook = new XlsxWorkbook(destination);
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
