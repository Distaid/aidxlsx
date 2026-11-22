using System.Globalization;
using NUnit.Framework;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.ConditionalFormats;

[TestFixture]
public class ConditionalFormatTests
{
    [Test]
    public void MaintainConditionalFormattingOrder()
    {
        // The input file contains duplicates of same dxf and thus the output also contains some duplicate dxf.
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\ConditionalFormattingOrder\inputfile.xlsx",
            @"Other\StyleReferenceFiles\ConditionalFormattingOrder\ConditionalFormattingOrder.xlsx");
    }

    [Test]
    public void DataBarKeepsItsNegativeColorAndIdOnRepeatedSave()
    {
        using var first = new MemoryStream();
        using (var workbook = new XlsxWorkbook())
        {
            var sheet = workbook.AddWorksheet("Bars");
            sheet.Range("B2:B5").AddConditionalFormat()
                .DataBar(XlsxColor.Blue, XlsxColor.Red, showBarOnly: true).LowestValue().HighestValue();
            workbook.SaveAs(first, validate: true);
        }

        using var second = new MemoryStream();
        Guid originalId;
        using (var workbook = new XlsxWorkbook(first))
        {
            originalId = ((XlsxConditionalFormat)workbook.Worksheet("Bars").ConditionalFormats.Single()).Id;
            workbook.SaveAs(second, validate: true);
        }

        using var loaded = new XlsxWorkbook(second);
        var format = (XlsxConditionalFormat)loaded.Worksheet("Bars").ConditionalFormats.Single();
        Assert.That(format.Id, Is.EqualTo(originalId));
        Assert.That(format.Colors.Values, Is.EqualTo(new[] { XlsxColor.Blue, XlsxColor.Red }));
        Assert.That(format.ShowBarOnly, Is.True);
    }

    [TestCase(true, 7)]
    [TestCase(false, 8)]
    public void SaveOptionAffectsConsolidationConditionalFormatRanges(bool consolidateConditionalFormatRanges, int expectedCount)
    {
        var options = new SaveOptions
        {
            ConsolidateConditionalFormatRanges = consolidateConditionalFormatRanges
        };

        var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet("Sheet");

        ws.Range("D2:D3").AddConditionalFormat().DataBar(XlsxColor.Red).LowestValue().HighestValue();
        ws.Range("B2:B3").AddConditionalFormat().DataBar(XlsxColor.Red).LowestValue().HighestValue();
        ws.Range("E2:E6").AddConditionalFormat().ColorScale().LowestValue(XlsxColor.Red).HighestValue(XlsxColor.Blue);
        ws.Range("F2:F6").AddConditionalFormat().ColorScale().LowestValue(XlsxColor.Red).HighestValue(XlsxColor.Blue);
        ws.Range("G2:G7").AddConditionalFormat().WhenIsUnique().Fill.SetBackgroundColor(XlsxColor.Blue);
        ws.Range("H2:H7").AddConditionalFormat().WhenIsUnique().Fill.SetBackgroundColor(XlsxColor.Blue);
        ws.Range("I2:I6").AddConditionalFormat().WhenContains("test");
        ws.Range("J2:J6").AddConditionalFormat().WhenContains("test");
        using (var ms = new MemoryStream())
        {
            wb.SaveAs(ms, options);
            var wb_saved = new XlsxWorkbook(ms);
            Assert.AreEqual(expectedCount, wb_saved.Worksheet("Sheet").ConditionalFormats.Count());
        }
    }

    [TestCase(true, 1)]
    [TestCase(false, 2)]
    public void SaveOptionAffectsConsolidationDataValidationRanges(bool consolidateDataValidationRanges, int expectedCount)
    {
        var options = new SaveOptions
        {
            ConsolidateDataValidationRanges = consolidateDataValidationRanges
        };

        var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Range("C2:C5").CreateDataValidation().Decimal.Between(1, 5);
        ws.Range("D2:D5").CreateDataValidation().Decimal.Between(1, 5);

        using (var ms = new MemoryStream())
        {
            wb.SaveAs(ms, options);
            var wb_saved = new XlsxWorkbook(ms);
            Assert.AreEqual(expectedCount, wb_saved.Worksheet("Sheet").DataValidations.Count());
        }
    }

    [TestCase("en-US")]
    [TestCase("fr-FR")]
    [TestCase("ru-RU")]
    public void SaveConditionalFormat_CultureIndependent(string culture)
    {
        using (var ms = new MemoryStream())
        {
            var expectedValue = 1.5;
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            using (var wb = new XlsxWorkbook())
            {
                var ws = wb.AddWorksheet();
                var i = 1;
                ws.Cell(i++, 1).AddConditionalFormat().WhenEquals(expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);
                ws.Cell(i++, 1).AddConditionalFormat().WhenNotEquals(expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);
                ws.Cell(i++, 1).AddConditionalFormat().WhenGreaterThan(expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);
                ws.Cell(i++, 1).AddConditionalFormat().WhenLessThan(expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);
                ws.Cell(i++, 1).AddConditionalFormat().WhenEqualOrGreaterThan(expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);
                ws.Cell(i++, 1).AddConditionalFormat().WhenEqualOrLessThan(expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);
                ws.Cell(i++, 1).AddConditionalFormat().WhenBetween(expectedValue, expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);
                ws.Cell(i++, 1).AddConditionalFormat().WhenNotBetween(expectedValue, expectedValue).Fill.SetBackgroundColor(XlsxColor.Red);

                wb.SaveAs(ms);
            }

            using (var wb = new XlsxWorkbook(ms))
            {
                var ws = wb.Worksheets.First();

                var conditionalFormatValues = ws.ConditionalFormats
                    .SelectMany(cf => cf.Values.Values)
                    .Select(v => v.Value)
                    .Distinct();

                Assert.AreEqual(1, conditionalFormatValues.Count());
                Assert.AreEqual("1.5", conditionalFormatValues.Single());
            }
        }
    }

    [Test]
    public void CellIs_type_reads_only_required_formula_arguments()
    {
        // The CellIs uses formula tags as arguments. Some producers generate extra empty
        // formula tags and Xlsx should be able to load CellIs conditional formatting
        // with such extra tags without an exception. The test file has been modified to
        // include extra formula tags and test checks that extra tags are ignored.
        TestHelper.LoadAndAssert((_, ws) =>
        {
            AssertFormulaArgs(ws, XlsxCFOperator.Between, "$D$2", "$E$2");
            AssertFormulaArgs(ws, XlsxCFOperator.NotBetween, "$D$3", "$E$3");
            AssertFormulaArgs(ws, XlsxCFOperator.GreaterThan, "$D$4");
            AssertFormulaArgs(ws, XlsxCFOperator.LessThan, "$D$5");
            AssertFormulaArgs(ws, XlsxCFOperator.Equal, "$D$6");
        }, @"Other\ConditionalFormats\Extra_formulas_CellIs_type.xlsx");

        static void AssertFormulaArgs(IXlsxWorksheet ws, XlsxCFOperator cfOperator, params string[] expectedFormulas)
        {
            var cf = ws.ConditionalFormats.Single(cf => cf.ConditionalFormatType == XlsxConditionalFormatType.CellIs && cf.Operator == cfOperator);
            Assert.AreEqual(expectedFormulas.Length, cf.Values.Count);
            CollectionAssert.AreEqual(expectedFormulas, cf.Values.Select(v => v.Value.Value));
        }
    }

    [Test]
    public void Expression_type_skips_empty_formula_tags()
    {
        // The Expression uses formula tag as arguments. Some producers generate extra empty
        // formula tags and Xlsx should be able to load Expression conditional formatting
        // with such extra tags without an exception. The test file has been modified to
        // include extra formula tags and test checks that extra tags are ignored.
        TestHelper.LoadAndAssert((_, ws) =>
        {
            AssertFormulaArgs(ws, "A1:A1", "$C$1=5");
            AssertFormulaArgs(ws, "A2:A2", "$C$2=4");
        }, @"Other\ConditionalFormats\Extra_formulas_Expression_type.xlsx");

        static void AssertFormulaArgs(IXlsxWorksheet ws, string range, string expectedFormula)
        {
            var cf = ws.ConditionalFormats.Single(cf => cf.ConditionalFormatType == XlsxConditionalFormatType.Expression && cf.Range.RangeAddress.ToString() == range);
            Assert.AreEqual(1, cf.Values.Count);
            CollectionAssert.AreEqual(expectedFormula, cf.Values[1].Value);
        }
    }

    [Test]
    public void Range_setter_throws_on_range_from_different_worksheet()
    {
        using var wb = new XlsxWorkbook();
        var ws1 = wb.AddWorksheet();
        var cf = ws1.Range("A1").AddConditionalFormat();
        var ws2 = wb.AddWorksheet();
        var differentSheetRange = ws2.Range("B5");

        Assert.Throws<ArgumentException>(() => cf.Range = differentSheetRange);
    }

    [Test]
    public void Ranges_setter_throws_on_range_from_different_worksheet()
    {
        using var wb = new XlsxWorkbook();
        var ws1 = wb.AddWorksheet();
        var cf = ws1.Range("A1").AddConditionalFormat();

        var ranges = ws1.Ranges("B1");
        var differentSheetRange = wb.AddWorksheet().Range("C1");
        ranges.Add(differentSheetRange);

        Assert.That(() => cf.Ranges = ranges, Throws.TypeOf<ArgumentException>().With.Message.Contains("must be from worksheet"));
    }

    [Test]
    public void Ranges_setter_throws_on_empty_ranges()
    {
        using var wb = new XlsxWorkbook();
        var ws1 = wb.AddWorksheet();
        var cf = ws1.Range("A1").AddConditionalFormat();

        var emptyRanges = ws1.Ranges("");

        Assert.That(() => cf.Ranges = emptyRanges, Throws.TypeOf<ArgumentException>().With.Message.Contains("empty"));
    }

    [Test]
    public void Text_formats_are_loaded_as_value()
    {
        // Issue #2690: Ensure that CF for text are loaded as a value and not as a formula.
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.AddConditionalFormat().WhenContains("Contains")
                    .Fill.SetBackgroundColor(XlsxColor.Red);

                ws.AddConditionalFormat().WhenNotContains("Not Contains")
                    .Fill.SetBackgroundColor(XlsxColor.Green);

                ws.AddConditionalFormat().WhenStartsWith("Starts With")
                    .Fill.SetBackgroundColor(XlsxColor.Blue);

                ws.AddConditionalFormat().WhenEndsWith("Ends With")
                    .Fill.SetBackgroundColor(XlsxColor.Black);
            },
            (_, ws) =>
            {
                AssertTextCfValue(ws, XlsxConditionalFormatType.ContainsText, "Contains");
                AssertTextCfValue(ws, XlsxConditionalFormatType.NotContainsText, "Not Contains");
                AssertTextCfValue(ws, XlsxConditionalFormatType.StartsWith, "Starts With");
                AssertTextCfValue(ws, XlsxConditionalFormatType.EndsWith, "Ends With");
            });
        return;

        static void AssertTextCfValue(IXlsxWorksheet ws, XlsxConditionalFormatType type, string text)
        {
            var cf = ws.ConditionalFormats.Single(x => x.ConditionalFormatType == type);
            Assert.AreEqual(1, cf.Values.Count);
            Assert.AreEqual(text, cf.Values[1].Value);
            Assert.IsFalse(cf.Values[1].IsFormula);
        }
    }
}
