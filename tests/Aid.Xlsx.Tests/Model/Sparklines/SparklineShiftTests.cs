using NUnit.Framework;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Sparklines;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Model.Sparklines;

[TestFixture]
public class SparklineShiftTests
{
    [Test]
    public void SparklineAreShiftedOnColumnInsert()
    {
        AssertSparklinePosition("D2", ws => ws.Column("C").InsertColumnsAfter(2), "F2");
    }

    [Test]
    public void SparklineAreShiftedOnColumnDelete()
    {
        AssertSparklinePosition("F2", ws => ws.Column("C").Delete(), "E2");
    }

    [Test]
    public void SparklineColumnShiftedOutOfSheetAreRemoved()
    {
        AssertSparklinePosition("XFD1", ws => ws.Column("C").InsertColumnsAfter(1), null);
    }

    [Test]
    public void SparklineAreShiftedOnRowInsert()
    {
        AssertSparklinePosition("B3", ws => ws.Row(2).InsertRowsBelow(3), "B6");
    }

    [Test]
    public void SparklineAreShiftedOnRowDelete()
    {
        AssertSparklinePosition("F8", ws => ws.Rows(4, 6).Delete(), "F5");
    }

    [Test]
    public void SparklineRowShiftedOutOfSheetAreRemoved()
    {
        AssertSparklinePosition($"A{XlsxHelper.MaxRowNumber}", ws => ws.Row(2).InsertRowsBelow(1), null);
    }

    private static void AssertSparklinePosition(string sparklineAddress, Action<IXlsxWorksheet> insertAction, string expectedAddress)
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("C2").Value = 2;
        var sparklineGroup = ws.SparklineGroups.Add(sparklineAddress, "B2:C2");
        insertAction(ws);
        Assert.AreEqual(expectedAddress, sparklineGroup.SingleOrDefault()?.Location.Address.ToString());
        if (expectedAddress is null)
            Assert.IsEmpty(sparklineGroup);
    }
}
