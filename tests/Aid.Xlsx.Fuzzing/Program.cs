using System.Text;
using System.Xml;
using Aid.Xlsx.Opc;
using Aid.Xlsx.Workbooks;
using SharpFuzz;

const int maxInput = 64 * 1024;
const string spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const string worksheetPart = "/xl/worksheets/sheet1.xml";
const string stylesPart = "/xl/styles.xml";

if (args is ["--seeds", var destination])
{
    WriteSeeds(destination);
    return 0;
}

if (args.Length != 1 || args[0] is not ("package" or "worksheet" or "styles" or "mce"))
{
    Console.Error.WriteLine("Select package, worksheet, styles, or mce, or use --seeds <directory>.");
    return 2;
}

var mode = args[0];
byte[]? template = null;
Fuzzer.Run(stream =>
{
    var bytes = new byte[maxInput + 1];
    var length = stream.ReadAtLeast(bytes, bytes.Length, throwOnEndOfStream: false);
    if (length > maxInput)
        return;

    using var input = new MemoryStream(bytes, 0, length, writable: false);
    try
    {
        switch (mode)
        {
            case "package":
                using (var package = OpcPackage.Open(input, new OpcOpenOptions
                {
                    MaxEntries = 128,
                    MaxUncompressedBytes = 512 * 1024,
                    MaxMetadataCharacters = maxInput
                }))
                {
                    foreach (var part in package.Parts.Values)
                    {
                        using var data = part.OpenRead();
                        data.CopyTo(Stream.Null);
                    }
                }
                break;
            case "worksheet":
            case "styles":
                using (var original = new MemoryStream(template ??= CreateWorkbook(), writable: false))
                using (var package = OpcPackage.Open(original))
                using (var updated = new MemoryStream())
                {
                    package.GetPart(mode == "worksheet" ? worksheetPart : stylesPart).ReplaceContent(bytes.AsSpan(0, length));
                    package.Save(updated);
                    updated.Position = 0;
                    using var workbook = new XlsxWorkbook(updated);
                }
                break;
            case "mce":
                using (var xml = XmlReader.Create(input, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    MaxCharactersInDocument = maxInput
                }))
                using (var mce = new MceXmlReader(xml, new MceSettings { ApplicationConfiguration = [spreadsheetNamespace] }))
                {
                    while (mce.Read())
                    {
                    }
                }
                break;
        }
    }
    catch (Exception exception) when (exception is InvalidDataException or XmlException or PartStructureException)
    {
        // Malformed containers or XML are expected to be rejected by the parsers.
    }
});
return 0;

static byte[] CreateWorkbook()
{
    using var workbook = new XlsxWorkbook();
    var sheet = workbook.AddWorksheet("Seed");
    sheet.Cell("A1").Value = 20;
    sheet.Cell("A2").Value = "A text cell";
    sheet.Cell("B1").FormulaA1 = "A1+22";
    sheet.Cell("A1").Style.Font.Bold = true;
    using var file = new MemoryStream();
    workbook.SaveAs(file, new SaveOptions { ValidatePackage = true, EvaluateFormulasBeforeSaving = true });
    return file.ToArray();
}

static void WriteSeeds(string destination)
{
    var workbook = CreateWorkbook();
    WriteSeed("package", workbook);
    using var stream = new MemoryStream(workbook, writable: false);
    using var package = OpcPackage.Open(stream);
    WritePart("worksheet", worksheetPart);
    WritePart("styles", stylesPart);
    WriteSeed("mce", Encoding.UTF8.GetBytes($"""
        <worksheet xmlns="{spreadsheetNamespace}" xmlns:s="{spreadsheetNamespace}"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:future" mc:Ignorable="f">
          <mc:AlternateContent>
            <mc:Choice Requires="s"><sheetData/></mc:Choice>
            <mc:Fallback><f:sheetData/></mc:Fallback>
          </mc:AlternateContent>
        </worksheet>
        """));

    void WritePart(string name, string partName)
    {
        using var part = package.GetPart(partName).OpenRead();
        using var buffer = new MemoryStream();
        part.CopyTo(buffer);
        WriteSeed(name, buffer.ToArray());
    }

    void WriteSeed(string mode, byte[] seed)
    {
        var directory = Path.Combine(destination, mode);
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "seed"), seed);
    }
}
