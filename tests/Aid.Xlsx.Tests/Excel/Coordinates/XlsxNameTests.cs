using Aid.Xlsx.Coordinates;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.Coordinates
{
    [TestFixture]
    public class XlsxNameTests
    {
        [Test]
        public void Workbook_scoped_name_is_compared_case_insensitive()
        {
            var lowerCase = new XlsxName("name");
            var upperCase = new XlsxName("NAME");

            Assert.AreEqual(lowerCase, upperCase);
            Assert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

            Assert.AreNotEqual(lowerCase, new XlsxName("different_name"));
        }

        [Test]
        public void Sheet_scoped_name_is_compared_case_insensitive()
        {
            var lowerCase = new XlsxName("sheet", "name");
            var upperCase = new XlsxName("SHEET", "NAME");

            Assert.AreEqual(lowerCase, upperCase);
            Assert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

            Assert.AreNotEqual(lowerCase, new XlsxName("Different sheet", "name"));
            Assert.AreNotEqual(lowerCase, new XlsxName("sheet", "different_name"));
        }
    }
}
