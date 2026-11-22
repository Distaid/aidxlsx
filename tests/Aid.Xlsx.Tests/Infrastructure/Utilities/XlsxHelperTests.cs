using NUnit.Framework;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Tests.Infrastructure.Utilities;

[TestFixture]
public class XlsxHelperTests
{
    [Test]
    public void IsValidColumnTest()
    {
        Assert.AreEqual(false, XlsxHelper.IsValidColumn(""));
        Assert.AreEqual(false, XlsxHelper.IsValidColumn("1"));
        Assert.AreEqual(false, XlsxHelper.IsValidColumn("A1"));
        Assert.AreEqual(false, XlsxHelper.IsValidColumn("AA1"));
        Assert.AreEqual(true, XlsxHelper.IsValidColumn("A"));
        Assert.AreEqual(true, XlsxHelper.IsValidColumn("AA"));
        Assert.AreEqual(true, XlsxHelper.IsValidColumn("AAA"));
        Assert.AreEqual(true, XlsxHelper.IsValidColumn("Z"));
        Assert.AreEqual(true, XlsxHelper.IsValidColumn("ZZ"));
        Assert.AreEqual(true, XlsxHelper.IsValidColumn("XFD"));
        Assert.AreEqual(false, XlsxHelper.IsValidColumn("ZAA"));
        Assert.AreEqual(false, XlsxHelper.IsValidColumn("XZA"));
        Assert.AreEqual(false, XlsxHelper.IsValidColumn("XFZ"));
    }

    [Test]
    public void ReplaceRelative1()
    {
        string result = XlsxHelper.ReplaceRelative("A1", 2, "B");
        Assert.AreEqual("B2", result);
    }

    [Test]
    public void ReplaceRelative2()
    {
        string result = XlsxHelper.ReplaceRelative("$A1", 2, "B");
        Assert.AreEqual("$A2", result);
    }

    [Test]
    public void ReplaceRelative3()
    {
        string result = XlsxHelper.ReplaceRelative("A$1", 2, "B");
        Assert.AreEqual("B$1", result);
    }

    [Test]
    public void ReplaceRelative4()
    {
        string result = XlsxHelper.ReplaceRelative("$A$1", 2, "B");
        Assert.AreEqual("$A$1", result);
    }

    [Test]
    public void ReplaceRelative5()
    {
        string result = XlsxHelper.ReplaceRelative("1:1", 2, "B");
        Assert.AreEqual("2:2", result);
    }

    [Test]
    public void ReplaceRelative6()
    {
        string result = XlsxHelper.ReplaceRelative("$1:1", 2, "B");
        Assert.AreEqual("$1:2", result);
    }

    [Test]
    public void ReplaceRelative7()
    {
        string result = XlsxHelper.ReplaceRelative("1:$1", 2, "B");
        Assert.AreEqual("2:$1", result);
    }

    [Test]
    public void ReplaceRelative8()
    {
        string result = XlsxHelper.ReplaceRelative("$1:$1", 2, "B");
        Assert.AreEqual("$1:$1", result);
    }

    [Test]
    public void ReplaceRelative9()
    {
        string result = XlsxHelper.ReplaceRelative("A:A", 2, "B");
        Assert.AreEqual("B:B", result);
    }

    [Test]
    public void ReplaceRelativeA()
    {
        string result = XlsxHelper.ReplaceRelative("$A:A", 2, "B");
        Assert.AreEqual("$A:B", result);
    }

    [Test]
    public void ReplaceRelativeB()
    {
        string result = XlsxHelper.ReplaceRelative("A:$A", 2, "B");
        Assert.AreEqual("B:$A", result);
    }

    [Test]
    public void ReplaceRelativeC()
    {
        string result = XlsxHelper.ReplaceRelative("$A:$A", 2, "B");
        Assert.AreEqual("$A:$A", result);
    }

    [TestCase("Sheet1", "Sheet1")]
    [TestCase("O'Brien's sales", "O'Brien's sales")]
    [TestCase(" data # ", " data # ")]
    [TestCase("data $1.00", "data $1.00")]
    [TestCase("data ", "data?")]
    [TestCase("abc def", "abc/def")]
    [TestCase("data 0 ", "data[0]")]
    [TestCase("data ", "data*")]
    [TestCase("abc def", "abc\\def")]
    [TestCase(" data", "'data")]
    [TestCase("data ", "data'")]
    [TestCase("d'at'a", "d'at'a")]
    [TestCase("sheet a4", "sheet:a4")]
    [TestCase("null", null)]
    [TestCase("empty", "")]
    [TestCase("1234567890123456789012345678901", "1234567890123456789012345678901TOOLONG")]
    public void CreateSafeSheetNames(string expected, string input)
    {
        var actual = XlsxHelper.CreateSafeSheetName(input);
        Assert.AreEqual(expected, actual);
    }

    [TestCase("Sheet1", ExpectedResult = "Sheet1")]
    [TestCase("O'Brien's sales", ExpectedResult = "O'Brien's sales")]
    [TestCase(" data # ", ExpectedResult = " data # ")]
    [TestCase("data $1.00", ExpectedResult = "data $1.00")]
    [TestCase("data?", ExpectedResult = "data_")]
    [TestCase("abc/def", ExpectedResult = "abc_def")]
    [TestCase("data[0]", ExpectedResult = "data_0_")]
    [TestCase("data*", ExpectedResult = "data_")]
    [TestCase("abc\\def", ExpectedResult = "abc_def")]
    [TestCase("'data", ExpectedResult = "_data")]
    [TestCase("data'", ExpectedResult = "data_")]
    [TestCase("d'at'a", ExpectedResult = "d'at'a")]
    [TestCase("sheet:a4", ExpectedResult = "sheet_a4")]
    [TestCase(null, ExpectedResult = "null")]
    [TestCase("", ExpectedResult = "empty")]
    [TestCase("1234567890123456789012345678901TOOLONG", ExpectedResult = "1234567890123456789012345678901")]
    public string CreateSafeSheetNamesWithUnderscore(string input)
    {
        return XlsxHelper.CreateSafeSheetName(input, replaceChar: '_');
    }

    [Test]
    public void CreateSafeSheetNamesInvalidReplacementChar()
    {
        Assert.Throws<ArgumentException>(() => XlsxHelper.CreateSafeSheetName("abc\\def", replaceChar: ':'));
    }
}
