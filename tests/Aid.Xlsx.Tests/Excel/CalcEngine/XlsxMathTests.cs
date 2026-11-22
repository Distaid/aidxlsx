using Aid.Xlsx.CalcEngine.Functions;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class XlsxMathTests
    {
        [Test]
        public void IsEven()
        {
            Assert.IsTrue(XlsxMath.IsEven(2));
            Assert.IsFalse(XlsxMath.IsEven(3));
        }

        [Test]
        public void IsOdd()
        {
            Assert.IsTrue(XlsxMath.IsOdd(3));
            Assert.IsFalse(XlsxMath.IsOdd(2));
        }
    }
}
