using Aid.Xlsx.Tests.Model.Styles;
using NUnit.Framework;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.Styles;

[TestFixture]
public class XlsxFillTests
{
    [Test]
    [TestCaseSource(nameof(FillApiSetters))]
    public void Fill_property_can_be_individually_set(FormatTestCase<IXlsxFill> testCase)
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();

        // Set two-color pattern, so setting individual property doesn't trigger special logic
        var cellFormat = ws.Cell("B2").Style
                .Fill.SetPatternType(XlsxFillPatternValues.LightGrid)
                .Fill.SetBackgroundColor(XlsxColor.Aqua)
                .Fill.SetPatternColor(XlsxColor.Lemon);

        foreach (var testValue in testCase.Values)
        {
            testCase.SetPropertyValue(cellFormat.Fill, testValue);
            var setValue = testCase.GetPropertyValue(cellFormat.Fill);
            Assert.AreEqual(testValue, setValue);
        }
    }

    private static IEnumerable<FormatTestCase<IXlsxFill>> FillApiSetters()
    {
        var patternValues = EnumPolyfill.GetValues<XlsxFillPatternValues>();
        yield return FormatTestCase<IXlsxFill>.ForFill(fill => fill.PatternType, (fill, value) => fill.PatternType = value, patternValues);
        yield return FormatTestCase<IXlsxFill>.ForFill(fill => fill.PatternType, (fill, value) => fill.SetPatternType(value), patternValues);

        var colors = new[] { XlsxColor.Black, XlsxColor.Red, XlsxColor.Automatic, XlsxColor.Transparent };
        yield return FormatTestCase<IXlsxFill>.ForFill(fill => fill.BackgroundColor, (fill, value) => fill.BackgroundColor = value, colors);
        yield return FormatTestCase<IXlsxFill>.ForFill(fill => fill.BackgroundColor, (fill, value) => fill.SetBackgroundColor(value), colors);

        yield return FormatTestCase<IXlsxFill>.ForFill(fill => fill.PatternColor, (fill, value) => fill.PatternColor = value, colors);
        yield return FormatTestCase<IXlsxFill>.ForFill(fill => fill.PatternColor, (fill, value) => fill.SetPatternColor(value), colors);
    }

    [Test]
    public void BackgroundColor_keeps_pattern_on_two_color_patterns()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var fill = ws.Cell("A1").Style.Fill;
        fill.PatternType = XlsxFillPatternValues.LightGrid;
        Assert.AreEqual(XlsxFillPatternValues.LightGrid, fill.PatternType);

        fill.BackgroundColor = XlsxColor.Blue;

        Assert.AreEqual(XlsxFillPatternValues.LightGrid, fill.PatternType);
    }

    [Test]
    public void BackgroundColor_sets_pattern_to_solid_when_original_pattern_was_none()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var fill = ws.Cell("A1").Style.Fill;
        Assert.AreEqual(XlsxFillPatternValues.None, fill.PatternType);

        fill.BackgroundColor = XlsxColor.Blue;

        Assert.AreEqual(XlsxFillPatternValues.Solid, fill.PatternType);
    }

    [Test]
    public void BackgroundColor_set_to_transparent_color_sets_pattern_to_none()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        var fill = ws.Cell("A1").Style.Fill;
        fill.BackgroundColor = XlsxColor.Red;
        Assert.AreEqual(XlsxFillPatternValues.Solid, fill.PatternType);

        fill.BackgroundColor = XlsxColor.Automatic;

        Assert.AreEqual(XlsxFillPatternValues.None, fill.PatternType);
    }

    [Test]
    public void BackgroundPatternEqualCheck()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();

        var fill1 = ws.Cell("A1").Style.Fill;
        fill1.BackgroundColor = XlsxColor.Blue;
        var fill2 = ws.Cell("A2").Style.Fill;
        fill2.BackgroundColor = XlsxColor.Blue;

        Assert.IsTrue(fill1.Equals(fill2));
        Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
    }

    [Test]
    public void BackgroundPatternNotEqualCheck()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();

        var fill1 = ws.Cell("A1").Style.Fill;
        fill1.PatternType = XlsxFillPatternValues.Solid;
        fill1.BackgroundColor = XlsxColor.Blue;

        var fill2 = ws.Cell("A2").Style.Fill;
        fill2.PatternType = XlsxFillPatternValues.Solid;
        fill2.BackgroundColor = XlsxColor.Red;

        Assert.IsFalse(fill1.Equals(fill2));
    }

    [Test]
    public void FillsWithStyleNoneAreEqualToFillWithAutomaticColor()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();

        var fill1 = ws.Cell("A1").Style.Fill.SetBackgroundColor(XlsxColor.ElectricUltramarine).Fill.SetPatternType(XlsxFillPatternValues.None).Fill;
        var fill2 = ws.Cell("A2").Style.Fill.SetBackgroundColor(XlsxColor.EtonBlue).Fill.SetPatternType(XlsxFillPatternValues.None).Fill;
        var fill3 = ws.Cell("A3").Style.Fill.SetBackgroundColor(XlsxColor.Automatic).Fill;

        Assert.IsTrue(fill1.Equals(fill2));
        Assert.IsTrue(fill1.Equals(fill3));
        Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
        Assert.AreEqual(fill1.GetHashCode(), fill3.GetHashCode());
    }

    [Test]
    public void SolidFillsWithDifferentPatternColorEqual()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();

        var fill1 = ws.Cell("A1").Style.Fill;
        fill1.PatternType = XlsxFillPatternValues.Solid;
        fill1.BackgroundColor = XlsxColor.Red;
        fill1.PatternColor = XlsxColor.Blue;

        var fill2 = ws.Cell("A2").Style.Fill;
        fill2.PatternType = XlsxFillPatternValues.Solid;
        fill2.BackgroundColor = XlsxColor.Red;
        fill2.PatternColor = XlsxColor.Green;

        Assert.IsTrue(fill1.Equals(fill2));
        Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
    }

    [Test]
    public void BackgroundWithConditionalFormat()
    {
        var workbook = new XlsxWorkbook();
        var worksheet = workbook.AddWorksheet("Test");
        worksheet.Cell(2, 2).SetValue("Text");
        var cf = worksheet.Cell(2, 2).AddConditionalFormat();
        var style = cf.WhenNotBlank();
        style = style
            .Border.SetOutsideBorder(XlsxBorderStyleValues.Thick)
            .Border.SetOutsideBorderColor(XlsxColor.Blue);

        Assert.AreEqual(style.Border.BottomBorder, XlsxBorderStyleValues.Thick);
        Assert.AreEqual(style.Border.TopBorder, XlsxBorderStyleValues.Thick);
        Assert.AreEqual(style.Border.LeftBorder, XlsxBorderStyleValues.Thick);
        Assert.AreEqual(style.Border.RightBorder, XlsxBorderStyleValues.Thick);

        Assert.AreEqual(style.Border.BottomBorderColor, XlsxColor.Blue);
        Assert.AreEqual(style.Border.TopBorderColor, XlsxColor.Blue);
        Assert.AreEqual(style.Border.LeftBorderColor, XlsxColor.Blue);
        Assert.AreEqual(style.Border.RightBorderColor, XlsxColor.Blue);
    }

    [Test]
    public void LoadAndSaveDxfBackgroundFill()
    {
        // The cells in test file have default format with a white background. Then, there are two CF:
        // * If value = 5, apply with automatic color background (=white)
        // * If value <> 5, apply red background
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\DxfBackgroundFill\inputfile.xlsx",
            @"Other\StyleReferenceFiles\DxfBackgroundFill\output.xlsx");
    }

    [Test]
    public void ReservedFills_ReplaceWithPredefinedValues()
    {
        // If attribute or whole predefined fill is missing from the file, save predefined values
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Input.xlsx",
            @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Output.xlsx");
    }

    [Test]
    public void ReservedFills_MoveFillsFromReservedPositions()
    {
        // If the input doesn't have expected fill values at the reserved position s0 and 1 (can only happen
        // for non-excel sources, excel always has correct values), put expected fill at 0 and 1, but save original
        // fills to different positions if they are used.
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Input.xlsx",
            @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Output.xlsx");
    }
}
