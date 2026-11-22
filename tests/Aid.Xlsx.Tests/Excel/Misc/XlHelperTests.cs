using NUnit.Framework;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Tests.Excel;

[TestFixture]
public class XlHelperTests
{
    private static void CheckColumnNumber(int column)
    {
        Assert.AreEqual(column, XlsxHelper.GetColumnNumberFromLetter(XlsxHelper.GetColumnLetterFromNumber(column)));
    }

    [Test]
    public void InvalidA1Addresses()
    {
        Assert.IsFalse(XlsxHelper.IsValidA1Address(""));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("A"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("a"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("-1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("AAAA1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("XFG1"));

        Assert.IsFalse(XlsxHelper.IsValidA1Address("@A1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("@AA1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("@AAA1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("[A1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("[AA1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("[AAA1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("{A1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("{AA1"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("{AAA1"));

        Assert.IsFalse(XlsxHelper.IsValidA1Address("A1@"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("AA1@"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("AAA1@"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("A1["));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("AA1["));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("AAA1["));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("A1{"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("AA1{"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("AAA1{"));

        Assert.IsFalse(XlsxHelper.IsValidA1Address("@A1@"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("@AA1@"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("@AAA1@"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("[A1["));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("[AA1["));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("[AAA1["));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("{A1{"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("{AA1{"));
        Assert.IsFalse(XlsxHelper.IsValidA1Address("{AAA1{"));
    }

    [Test]
    public void PlusAA1_Is_Not_an_address()
    {
        Assert.IsFalse(XlsxHelper.IsValidA1Address("+AA1"));
    }

    [Test]
    public void TestConvertColumnLetterToNumberAnd()
    {
        CheckColumnNumber(1);
        CheckColumnNumber(27);
        CheckColumnNumber(28);
        CheckColumnNumber(52);
        CheckColumnNumber(53);
        CheckColumnNumber(1000);
        CheckColumnNumber(1353);
    }

    [Test]
    public void ValidA1Addresses()
    {
        Assert.IsTrue(XlsxHelper.IsValidA1Address("A1"));
        Assert.IsTrue(XlsxHelper.IsValidA1Address("A" + XlsxHelper.MaxRowNumber));
        Assert.IsTrue(XlsxHelper.IsValidA1Address("Z1"));
        Assert.IsTrue(XlsxHelper.IsValidA1Address("Z" + XlsxHelper.MaxRowNumber));

        Assert.IsTrue(XlsxHelper.IsValidA1Address("AA1"));
        Assert.IsTrue(XlsxHelper.IsValidA1Address("AA" + XlsxHelper.MaxRowNumber));
        Assert.IsTrue(XlsxHelper.IsValidA1Address("ZZ1"));
        Assert.IsTrue(XlsxHelper.IsValidA1Address("ZZ" + XlsxHelper.MaxRowNumber));

        Assert.IsTrue(XlsxHelper.IsValidA1Address("AAA1"));
        Assert.IsTrue(XlsxHelper.IsValidA1Address("AAA" + XlsxHelper.MaxRowNumber));
        Assert.IsTrue(XlsxHelper.IsValidA1Address(XlsxHelper.MaxColumnLetter + "1"));
        Assert.IsTrue(XlsxHelper.IsValidA1Address(XlsxHelper.MaxColumnLetter + XlsxHelper.MaxRowNumber));
    }

    [Test]
    public void TestColumnLetterLookup()
    {
        var columnLetters = new List<String>();
        for (int c = 1; c <= XlsxHelper.MaxColumnNumber; c++)
        {
            var columnLetter = NaiveGetColumnLetterFromNumber(c);
            columnLetters.Add(columnLetter);

            Assert.AreEqual(columnLetter, XlsxHelper.GetColumnLetterFromNumber(c));
        }

        foreach (var cl in columnLetters)
        {
            var columnNumber = NaiveGetColumnNumberFromLetter(cl);
            Assert.AreEqual(columnNumber, XlsxHelper.GetColumnNumberFromLetter(cl));
        }
    }

    [TestCase("R")]
    [TestCase("C")]
    [TestCase("RC")]
    [TestCase("R111C222")]
    [TestCase("R[]C")]
    [TestCase("RC[]")]
    [TestCase("R[]C[]")]
    [TestCase("R[111]C222")]
    [TestCase("R111C[222]")]
    [TestCase("R[111]C[222]")]
    [TestCase("R[-111]C[-222]")]
    public void ValidRCAddresses(string address)
    {
        Assert.IsTrue(XlsxHelper.IsValidRCAddress(address));
    }

    [TestCase("RD")]
    [TestCase("CC")]
    [TestCase("R[-]C222")]
    [TestCase("R[]C[-]")]
    [TestCase("_R111C222")]
    public void InvalidRCAddresses(string address)
    {
        Assert.IsFalse(XlsxHelper.IsValidRCAddress(address));
    }

    #region Old XlsxHelper methods

    private static readonly string[] letters = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

    /// <summary>
    /// These used to be the methods in XlsxHelper, but were later changed
    /// We now use them as a check against the new methods
    /// Gets the column number of a given column letter.
    /// </summary>
    /// <param name="columnLetter"> The column letter to translate into a column number. </param>
    private static int NaiveGetColumnNumberFromLetter(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter)) throw new ArgumentNullException("columnLetter");

        int retVal;
        columnLetter = columnLetter.ToUpper();

        //Extra check because we allow users to pass row col positions in as strings
        if (columnLetter[0] <= '9')
        {
            retVal = Int32.Parse(columnLetter, XlsxHelper.NumberStyle, XlsxHelper.ParseCulture);
            return retVal;
        }

        int sum = 0;

        for (int i = 0; i < columnLetter.Length; i++)
        {
            sum *= 26;
            sum += (columnLetter[i] - 'A' + 1);
        }

        return sum;
    }

    /// <summary>
    /// Gets the column letter of a given column number.
    /// </summary>
    /// <param name="columnNumber">The column number to translate into a column letter.</param>
    /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
    private static string NaiveGetColumnLetterFromNumber(int columnNumber, bool trimToAllowed = false)
    {
        if (trimToAllowed) columnNumber = XlsxHelper.TrimColumnNumber(columnNumber);

        columnNumber--; // Adjust for start on column 1
        if (columnNumber <= 25)
        {
            return letters[columnNumber];
        }
        var firstPart = (columnNumber) / 26;
        var remainder = ((columnNumber) % 26) + 1;
        return NaiveGetColumnLetterFromNumber(firstPart) + NaiveGetColumnLetterFromNumber(remainder);
    }

    #endregion Old XlsxHelper methods
}
