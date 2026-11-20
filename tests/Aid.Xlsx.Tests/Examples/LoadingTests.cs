using ClosedXML.Examples;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Examples
{
    [TestFixture]
    public class LoadingTests
    {
        [Test]
        public void ChangingBasicTable()
        {
            TestHelper.RunTestExample<ChangingBasicTable>(@"Loading\ChangingBasicTable.xlsx");
        }
    }
}
