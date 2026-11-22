using System.IO.Packaging;
using System.Reflection;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NUnit.Framework;
using LoadOptions = Aid.Xlsx.Model.Workbooks.LoadOptions;
using Path = System.IO.Path;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.TestSupport;

internal static partial class TestHelper
{
    public static string CurrencySymbol
    {
        get { return Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol; }
    }

    //Note: Run example tests parameters
    public static string TestsOutputDirectory
    {
        get
        {
            return Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Generated");
        }
    }

    public const string ActualTestResultPostFix = "";
    public static readonly string ExampleTestsOutputDirectory = Path.Combine(TestsOutputDirectory, "Examples");

    private const bool CompareWithResources = true;

    private static readonly ResourceFileExtractor _extractor = new(Assembly.GetExecutingAssembly(), ".Resource.");

    public static void SaveWorkbook(XlsxWorkbook workbook, params string[] fileNameParts)
    {
        workbook.SaveAs(Path.Combine(new string[] { TestsOutputDirectory }.Concat(fileNameParts).ToArray()), true);
    }

    // Because different fonts are installed on Unix,
    // the columns widths after AdjustToContents() will
    // cause the tests to fail.
    // Therefore we ignore the width attribute when running on Unix
    public static bool StripColumnWidths { get { return IsRunningOnUnix; } }

    public static bool IsRunningOnUnix
    {
        get
        {
            int p = (int)Environment.OSVersion.Platform;
            return ((p == 4) || (p == 6) || (p == 128));
        }
    }

    /// <summary>
    /// Create a workbook and compare it with a saved resource.
    /// </summary>
    /// <param name="workbookGenerator">A function that gets an empty workbook and fills it with data.</param>
    /// <param name="referenceResource">Reference workbook saved in resources</param>
    /// <param name="evaluateFormulae">Should formulas of created workbook be evaluated and values saved?</param>
    /// <param name="validate">Should the created workbook be checked by the library and the Open XML SDK validator?</param>
    public static void CreateAndCompare(Action<XlsxWorkbook> workbookGenerator, string referenceResource, bool evaluateFormulae = false, bool validate = true)
    {
        CreateAndCompare(() =>
        {
            var wb = new XlsxWorkbook();
            workbookGenerator(wb);
            return wb;
        }, referenceResource, evaluateFormulae, validate);
    }

    public static void CreateAndCompare(Func<IXlsxWorkbook> workbookGenerator, string referenceResource, bool evaluateFormulae = false, bool validate = true)
    {
        Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

        string[] pathParts = referenceResource.Split(new char[] { '\\' });
        string filePath1 = Path.Combine(new List<string>() { TestsOutputDirectory }.Concat(pathParts).ToArray());

        var extension = Path.GetExtension(filePath1);
        var directory = Path.GetDirectoryName(filePath1);

        var fileName = Path.GetFileNameWithoutExtension(filePath1);
        fileName += ActualTestResultPostFix;
        fileName = Path.ChangeExtension(fileName, extension);

        var filePath2 = Path.Combine(directory, fileName);

        using (var wb = workbookGenerator.Invoke())
            wb.SaveAs(filePath2, validate, evaluateFormulae);

        if (validate)
        {
            using var document = SpreadsheetDocument.Open(filePath2, false);
            var errors = new OpenXmlValidator().Validate(document)
                .Select(error => $"Part {error.Part?.Uri}, path {error.Path?.XPath}: {error.Description}")
                .ToArray();
            Assert.That(errors, Is.Empty, string.Join(Environment.NewLine, errors));
        }

        if (CompareWithResources)
        {
            string resourcePath = referenceResource.Replace('\\', '.').TrimStart('.');
            using (var streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath))
            using (var streamActual = File.OpenRead(filePath2))
            {
                var success = ExcelDocsComparer.Compare(streamActual, streamExpected, out string message);
                var formattedMessage =
                    String.Format(
                        "Actual file '{0}' is different than the expected file '{1}'. The difference is: '{2}'",
                        filePath2, resourcePath, message);

                Assert.IsTrue(success, formattedMessage);
            }
        }
    }

    /// <summary>
    /// Load a file from the <paramref name="loadResourcePath"/>, modify it, save it through Aid.Xlsx
    /// and compare the saved file against the <paramref name="expectedOutputResourcePath"/>.
    /// </summary>
    /// <remarks>Useful for checking whether we can load data from Excel and save it while keeping various feature in the OpenXML intact.</remarks>
    public static void LoadModifyAndCompare(string loadResourcePath, Action<XlsxWorkbook> modify, string expectedOutputResourcePath, bool evaluateFormulae = false, bool validate = true)
    {
        using var stream = GetStreamFromResource(GetResourcePath(loadResourcePath));
        using var ms = new MemoryStream();
        CreateAndCompare(() =>
        {
            var wb = new XlsxWorkbook(stream);
            modify(wb);
            return wb;
        }, expectedOutputResourcePath, evaluateFormulae, validate);
    }

    /// <summary>
    /// Load a file from the <paramref name="loadResourcePath"/>, save it through Aid.Xlsx without modifications
    /// and compare the saved file against the <paramref name="expectedOutputResourcePath"/>.
    /// </summary>
    /// <remarks>Useful for checking whether we can load data from Excel and save it while keeping various feature in the OpenXML intact.</remarks>
    public static void LoadSaveAndCompare(string loadResourcePath, string expectedOutputResourcePath, bool evaluateFormulae = false, bool validate = true)
    {
        LoadModifyAndCompare(loadResourcePath, _ => { }, expectedOutputResourcePath, evaluateFormulae, validate);
    }

    /// <summary>
    /// A testing method to load a workbook from resource and assert the state of the loaded workbook.
    /// </summary>
    public static void LoadAndAssert(Action<XlsxWorkbook> assertWorkbook, string loadResourcePath, LoadOptions options = null)
    {
        using var stream = GetStreamFromResource(GetResourcePath(loadResourcePath));
        using var wb = new XlsxWorkbook(stream, options ?? new LoadOptions());

        assertWorkbook(wb);
    }

    /// <summary>
    /// A testing method to load a workbook with a single worksheet from resource and assert
    /// the state of the loaded workbook.
    /// </summary>
    public static void LoadAndAssert(Action<XlsxWorkbook, IXlsxWorksheet> assertWorksheet, string loadResourcePath, LoadOptions options = null)
    {
        LoadAndAssert(wb =>
        {
            var ws = wb.Worksheets.Single();
            assertWorksheet(wb, ws);
        }, loadResourcePath, options);
    }

    public static string GetResourcePath(string filePartName)
    {
        return filePartName.Replace('\\', '.').TrimStart('.');
    }

    public static Stream GetStreamFromResource(string resourcePath)
    {
        return _extractor.ReadFileFromResourceToStream(resourcePath);
    }

    public static void LoadFile(string filePartName)
    {
        IXlsxWorkbook wb;
        using (var stream = GetStreamFromResource(GetResourcePath(filePartName)))
        {
            Assert.DoesNotThrow(() => wb = new XlsxWorkbook(stream), "Unable to load resource {0}", filePartName);
        }
    }

    public static IEnumerable<String> ListResourceFiles(Func<String, Boolean> predicate = null)
    {
        return _extractor.GetFileNames(predicate);
    }

    /// <summary>
    /// A method for testing of a saving and loading capabilities of Aid.Xlsx. Use this
    /// method to check properties are correctly saved and loaded.
    /// </summary>
    /// <remarks>This method is specialized, so it only works on one sheet.</remarks>
    /// <param name="createWorksheet">
    /// Method to setup a worksheet that will be saved and the saved file will be compared to
    /// <paramref name="referenceResource"/>.
    /// </param>
    /// <param name="assertLoadedWorkbook">
    /// <paramref name="referenceResource"/> will be loaded and this method will check that it
    /// was loaded correctly (i.e. properties are what was set in <paramref name="createWorksheet"/>).
    /// </param>
    /// <param name="referenceResource">Saved reference file.</param>
    public static void CreateSaveLoadAssert(Action<XlsxWorkbook, IXlsxWorksheet> createWorksheet, Action<XlsxWorkbook, IXlsxWorksheet> assertLoadedWorkbook, string referenceResource)
    {
        CreateAndCompare(wb =>
        {
            var ws = wb.AddWorksheet();
            createWorksheet(wb, ws);
        }, referenceResource);
        LoadAndAssert(assertLoadedWorkbook, referenceResource);
    }

    /// <summary>
    /// Basically can survive through save and load cycle. Doesn't check against actual file.
    /// Useful for testing is internal structures are correctly initialized after load.
    /// </summary>
    /// <param name="createWorksheet">Code to create a workbook.</param>
    /// <param name="assertLoadedWorkbook">Method to assert that workbook was loaded correctly.</param>
    /// <param name="validate">Validate created workbook that it is a valid OOXML file by OpenXML SDK.</param>
    /// <param name="evaluateFormulas">Evaluate formulas during saving and save the evaluated results to the workbook file.</param>
    public static void CreateSaveLoadAssert(Action<XlsxWorkbook, IXlsxWorksheet> createWorksheet, Action<XlsxWorkbook, IXlsxWorksheet> assertLoadedWorkbook, bool validate = true, bool evaluateFormulas = false)
    {
        using var ms = new MemoryStream();
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet();
            createWorksheet(wb, ws);
            wb.SaveAs(ms, validate, evaluateFormulas);
        }

        using (var wb = new XlsxWorkbook(ms))
        {
            var ws = wb.Worksheets.Single();
            assertLoadedWorkbook(wb, ws);
        }
    }

    /// <summary>
    /// Test if some aspect of a workbook can survive through save and load cycle.
    /// </summary>
    public static void CreateSaveLoadAssert(Action<XlsxWorkbook> createWorksheet, Action<XlsxWorkbook> assertLoadedWorkbook, bool validate = true, bool evaluateFormulas = false)
    {
        using var ms = new MemoryStream();
        using (var wb = new XlsxWorkbook())
        {
            createWorksheet(wb);
            wb.SaveAs(ms, validate, evaluateFormulas);
        }

        ms.Position = 0;
        using (var wb = new XlsxWorkbook(ms))
        {
            assertLoadedWorkbook(wb);
        }
    }

    /// <summary>
    /// A method for testing the saving of a workbook. It loads and saves the workbook and uses
    /// assert methods for individual parts to check that the relevant parts were saved correctly.
    /// Unlike the <see cref="LoadSaveAndCompare"/>, this method is more resistant to changes
    /// in the saving code that are not directly related to the tested parts.
    /// </summary>
    internal static void LoadSaveAndAssert(string referenceResource, string part1, Action<XDocument> part1Assert, string part2, Action<XDocument> part2Assert)
    {
        LoadSaveAndAssert(referenceResource, [(part1, part1Assert), (part2, part2Assert)]);
    }

    private static void LoadSaveAndAssert(string loadResourcePath, (string PartPath, Action<XDocument> PartAssert)[] parts)
    {
        using var ms = new MemoryStream();
        using (var stream = GetStreamFromResource(GetResourcePath(loadResourcePath)))
        {
            using var wb = new XlsxWorkbook(stream);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var package = Package.Open(ms, FileMode.Open, FileAccess.Read);
        foreach (var (partPath, partAssert) in parts)
        {
            var part = package.GetPart(new Uri(partPath, UriKind.Relative));
            using var partStream = part.GetStream();
            var partXml = XDocument.Load(partStream);
            partAssert(partXml);
        }
    }
}
