using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Utils;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.Coordinates;

[TestOf(typeof(XlsxColumnArea))]
internal class XlsxColumnAreaTests
{
    [TestCase(null)]
    [TestCase("")]
    public void Ctor_sheet_must_be_valid(string invalidSheetName)
    {
        Assert.That(
            () => new XlsxColumnArea(invalidSheetName, 1),
            Throws.Exception.TypeOf<ArgumentException>());
    }

    [TestCase(-50)]
    [TestCase(0)]
    [TestCase(XlsxHelper.MaxColumnNumber + 1)]
    [TestCase(int.MaxValue)]
    public void Ctor_column_number_must_be_valid(int invalidColumnNumber)
    {
        Assert.That(
            () => new XlsxColumnArea("some sheet", invalidColumnNumber),
            Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
    }

    [TestCase("name", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 4, false)]
    [TestCase("some name", 1, "other name", 1, false)]
    public void Two_areas_are_compared_by_case_insensitive_sheet_name_and_column_number(string firstName, int firstColumn, string secondName, int secondColumn, bool areEqual)
    {
        var first = new XlsxColumnArea(firstName, firstColumn);
        var second = new XlsxColumnArea(secondName, secondColumn);
        Assert.That(first == second, Is.EqualTo(areEqual));
        Assert.That(first.GetHashCode() == second.GetHashCode(), Is.EqualTo(areEqual));
    }

    [Test]
    public void Area_property_returns_area_of_column()
    {
        var column = new XlsxColumnArea("name", 4);
        var columnArea = column.Area;
        Assert.AreEqual(columnArea, new SheetArea("name", new Area(1, 4, XlsxHelper.MaxRowNumber, 4)));
    }

    [TestCase("name", 4, "name!D:D")]
    [TestCase("some name", 26, "'some name'!Z:Z")]
    [TestCase("Joe's", 27, "'Joe''s'!AA:AA")]
    [TestCase("Joe", XlsxHelper.MaxColumnNumber, "Joe!XFD:XFD")]
    public void ToString_returns_readable_reference(string name, int columnNumber, string expected)
    {
        var columnArea = new XlsxColumnArea(name, columnNumber);
        Assert.AreEqual(expected, columnArea.ToString());
    }
}
