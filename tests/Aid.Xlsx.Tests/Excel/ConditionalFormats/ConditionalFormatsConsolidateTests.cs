using System.Diagnostics.CodeAnalysis;
using NUnit.Framework;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel.ConditionalFormats;

[TestFixture]
public class ConditionalFormatsConsolidateTests
{
    [Test]
    public void ConsecutivelyRowsConsolidateTest()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B2:C2").AddConditionalFormat());
        SetFormat1(ws.Range("B4:C4").AddConditionalFormat());
        SetFormat1(ws.Range("B3:C3").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(1, ws.ConditionalFormats.Count());
        var format = ws.ConditionalFormats.First();
        Assert.AreEqual("B2:C4", format.Range.RangeAddress.ToStringRelative());
        Assert.AreEqual("F2", format.Values.Values.First().Value);
    }

    [Test]
    public void ConsecutivelyColumnsConsolidateTest()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("D2:D3").AddConditionalFormat());
        SetFormat1(ws.Range("B2:B3").AddConditionalFormat());
        SetFormat1(ws.Range("C2:C3").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(1, ws.ConditionalFormats.Count());
        var format = ws.ConditionalFormats.First();
        Assert.AreEqual("B2:D3", format.Ranges.First().RangeAddress.ToStringRelative());
        Assert.AreEqual("F2", format.Values.Values.First().Value);
    }

    [Test]
    public void Contains1ConsolidateTest()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
        SetFormat1(ws.Range("C12:D12").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(1, ws.ConditionalFormats.Count());
        var format = ws.ConditionalFormats.First();
        Assert.AreEqual("B11:D12", format.Range.RangeAddress.ToStringRelative());
        Assert.AreEqual("F11", format.Values.Values.First().Value);
    }

    [Test]
    public void Contains2ConsolidateTest()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B14:C14").AddConditionalFormat());
        SetFormat1(ws.Range("B14:B14").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(1, ws.ConditionalFormats.Count());
        var format = ws.ConditionalFormats.First();
        Assert.AreEqual("B14:C14", format.Range.RangeAddress.ToStringRelative());
        Assert.AreEqual("F14", format.Values.Values.First().Value);
    }

    [Test]
    public void SuperimposedConsolidateTest()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B16:D18").AddConditionalFormat());
        SetFormat1(ws.Range("B18:D19").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(1, ws.ConditionalFormats.Count());
        var format = ws.ConditionalFormats.First();
        Assert.AreEqual("B16:D19", format.Range.RangeAddress.ToStringRelative());
        Assert.AreEqual("F16", format.Values.Values.First().Value);
    }

    [Test]
    public void DifferentFormatNoConsolidateTest()
    {
        var wb = new XlsxWorkbook();
        IXlsxWorksheet ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
        SetFormat2(ws.Range("C12:D12").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(2, ws.ConditionalFormats.Count());
    }

    [Test]
    public void ConsolidatePreservesPriorities()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.Worksheets.Add();

        // Format2 A1:A5 and A6:A10 can be consolidated without changing a priority if other CF rules.
        // Format1 A1:A5 and A6:A10 can't be consolidated, because there is a CF rule between them.
        SetFormat1(ws.Range("A1:A5").AddConditionalFormat());
        SetFormat2(ws.Range("A1:A5").AddConditionalFormat());
        SetFormat2(ws.Range("A6:A10").AddConditionalFormat());
        SetFormat1(ws.Range("A6:A10").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(3, ws.ConditionalFormats.Count());
        var cf1 = ws.ConditionalFormats.First();
        var cf2 = ws.ConditionalFormats.ElementAt(1);
        var cf3 = ws.ConditionalFormats.Last();
        Assert.That(cf1, Is.EqualTo(cf3).Using(new CfFormatComaparer()));
        Assert.That(cf1, Is.Not.EqualTo(cf2).Using(new CfFormatComaparer()));
    }

    [Test]
    public void ConsolidatePreservesPriorities2()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.Worksheets.Add();

        SetFormat1(ws.Range("A1:A1").AddConditionalFormat());
        SetFormat2(ws.Range("A2:A3").AddConditionalFormat());
        SetFormat1(ws.Range("A2:A6").AddConditionalFormat());
        SetFormat1(ws.Range("A7:A8").AddConditionalFormat());

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(3, ws.ConditionalFormats.Count());
        var cf1 = ws.ConditionalFormats.ElementAt(0);
        var cf2 = ws.ConditionalFormats.ElementAt(1);
        var cf3 = ws.ConditionalFormats.ElementAt(2);
        Assert.That(cf1, Is.EqualTo(cf3).Using(new CfFormatComaparer()));
        Assert.That(cf1, Is.Not.EqualTo(cf2).Using(new CfFormatComaparer()));
        Assert.IsTrue(ws.ConditionalFormats.All(cf => cf.Ranges.Count() == 1), "Number of ranges in consolidated conditional formats is expected to be 1");
        Assert.AreEqual("A1:A1", cf1.Ranges.Single().RangeAddress.ToString());
        Assert.AreEqual("A2:A3", cf2.Ranges.Single().RangeAddress.ToString());
        Assert.AreEqual("A2:A8", cf3.Ranges.Single().RangeAddress.ToString());
    }

    [Test]
    public void ConsolidateShiftsFormulaRelativelyToTopMostCell()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.Worksheets.Add();

        var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8");
        var cf = ranges.First().AddConditionalFormat();
        cf.Ranges = ranges;
        cf.Values.Add(new XlsxFormula("=A3=$D3"));
        cf.Style.Fill.SetBackgroundColor(XlsxColor.Red);
        ws.ConditionalFormats.Add(cf);

        ((XlsxConditionalFormats)ws.ConditionalFormats).Consolidate();

        Assert.AreEqual(1, ws.ConditionalFormats.Count());
        var consolidatedCf = ws.ConditionalFormats.Single();
        Assert.That(consolidatedCf, Is.EqualTo(cf).Using(new CfFormatComaparer()));
        Assert.AreEqual("A3:C8", consolidatedCf.Ranges.Single().RangeAddress.ToString());
        Assert.IsTrue(consolidatedCf.Values.Single().Value.IsFormula);
        Assert.AreEqual("A3=$D3", consolidatedCf.Values.Single().Value.Value);
    }

    [Test]
    public void ColorScaleComparing()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet");

            var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8");
            var cf1 = ranges.First().AddConditionalFormat();
            cf1.Ranges = ranges;
            cf1.ColorScale()
                .LowestValue(XlsxColor.Red)
                .HighestValue(XlsxColor.Green);

            var cf2 = ranges.First().AddConditionalFormat();
            cf2.Ranges = ranges;
            cf2.ColorScale()
                .LowestValue(XlsxColor.Red)
                .HighestValue(XlsxColor.Green);
            Assert.AreNotSame(cf1, cf2);
            Assert.True(XlsxConditionalFormat.NoRangeComparer.Equals((XlsxConditionalFormat)cf1, (XlsxConditionalFormat)cf2));
        }
    }

    private static void SetFormat1(IXlsxConditionalFormat format)
    {
        format.WhenEquals("=" + format.Range.FirstCell().CellRight(4).Address.ToStringRelative()).Fill.SetBackgroundColor(XlsxColor.Blue);
    }

    private static void SetFormat2(IXlsxConditionalFormat format)
    {
        format.WhenEquals(5).Fill.SetBackgroundColor(XlsxColor.AliceBlue);
    }

    private class CfFormatComaparer : IEqualityComparer<IXlsxConditionalFormat>
    {
        public bool Equals(IXlsxConditionalFormat x, IXlsxConditionalFormat y)
        {
            var lhs = (XlsxConditionalFormat)x;
            var rhs = (XlsxConditionalFormat)y;
            return lhs.FormatValue == rhs.FormatValue;
        }

        public int GetHashCode([DisallowNull] IXlsxConditionalFormat obj)
        {
            return HashCode.Combine(((XlsxConditionalFormat)obj).FormatValue);
        }
    }
}
