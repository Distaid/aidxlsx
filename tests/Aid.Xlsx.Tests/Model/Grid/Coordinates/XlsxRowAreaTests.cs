using NUnit.Framework;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Tests.Model.Grid.Coordinates;

[TestOf(typeof(XlsxRowArea))]
internal class XlsxRowAreaTests
{
    [TestCase(null)]
    [TestCase("")]
    public void Ctor_sheet_must_be_valid(string invalidSheetName)
    {
        Assert.That(
            () => new XlsxRowArea(invalidSheetName, 1),
            Throws.InstanceOf<ArgumentException>());
    }

    [TestCase(-50)]
    [TestCase(0)]
    [TestCase(XlsxHelper.MaxRowNumber + 1)]
    [TestCase(int.MaxValue)]
    public void Ctor_row_number_must_be_valid(int invalidRowNumber)
    {
        Assert.That(
            () => new XlsxRowArea("some sheet", invalidRowNumber),
            Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
    }

    [TestCase("name", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 4, false)]
    [TestCase("some name", 1, "other name", 1, false)]
    public void Two_areas_are_compared_by_case_insensitive_sheet_name_and_row_number(string firstName, int firstRow, string secondName, int secondRow, bool areEqual)
    {
        var first = new XlsxRowArea(firstName, firstRow);
        var second = new XlsxRowArea(secondName, secondRow);
        Assert.That(first == second, Is.EqualTo(areEqual));
        Assert.That(first.GetHashCode() == second.GetHashCode(), Is.EqualTo(areEqual));
    }

    [Test]
    public void Area_property_returns_area_of_row()
    {
        var row = new XlsxRowArea("name", 4);
        var rowArea = row.Area;
        Assert.AreEqual(rowArea, new SheetArea("name", new Area(4, XlsxHelper.MinColumnNumber, 4, XlsxHelper.MaxColumnNumber)));
    }

    [TestCase("name", 4, "name!4:4")]
    [TestCase("some name", 4, "'some name'!4:4")]
    [TestCase("Joe's", 4, "'Joe''s'!4:4")]
    [TestCase("Joe", XlsxHelper.MaxRowNumber, "Joe!1048576:1048576")]
    public void ToString_returns_readable_reference(string name, int rowNumber, string expected)
    {
        var rowArea = new XlsxRowArea(name, rowNumber);
        Assert.AreEqual(expected, rowArea.ToString());
    }
}
